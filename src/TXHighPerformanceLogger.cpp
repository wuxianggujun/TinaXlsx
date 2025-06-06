//
// @file TXHighPerformanceLogger.cpp
// @brief 🚀 高性能日志库实现
//

#include "TinaXlsx/TXHighPerformanceLogger.hpp"
#include <iostream>
#include <filesystem>

namespace TinaXlsx {

// ==================== TXConsoleSink 实现 ====================

TXConsoleSink::TXConsoleSink(bool colored) : colored_output_(colored) {}

void TXConsoleSink::write(const TXLogEntry& entry) {
    std::lock_guard<std::mutex> lock(write_mutex_);

    auto time_t = std::chrono::system_clock::to_time_t(entry.timestamp);
    auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(
        entry.timestamp.time_since_epoch()) % 1000;

    if (colored_output_) {
        std::cout << getColorCode(entry.level);
    }

    // 🚀 使用fmt高效格式化
    std::cout << fmt::format("[{:%Y-%m-%d %H:%M:%S}.{:03d}] [{}] [{}] {}\n",
                            fmt::localtime(time_t), ms.count(),
                            getLevelString(entry.level), entry.logger_name, entry.message);

    if (colored_output_) {
        std::cout << "\033[0m"; // 重置颜色
    }

    // 🚀 立即刷新确保输出
    std::cout.flush();
}

void TXConsoleSink::flush() {
    std::cout.flush();
}

std::string TXConsoleSink::getColorCode(TXLogLevel level) const {
    switch (level) {
        case TXLogLevel::TRACE: return "\033[37m";    // 白色
        case TXLogLevel::DEBUG: return "\033[36m";    // 青色
        case TXLogLevel::INFO: return "\033[32m";     // 绿色
        case TXLogLevel::WARN: return "\033[33m";     // 黄色
        case TXLogLevel::ERROR: return "\033[31m";    // 红色
        case TXLogLevel::CRITICAL: return "\033[35m"; // 紫色
        default: return "";
    }
}

std::string TXConsoleSink::getLevelString(TXLogLevel level) const {
    switch (level) {
        case TXLogLevel::TRACE: return "TRACE";
        case TXLogLevel::DEBUG: return "DEBUG";
        case TXLogLevel::INFO: return "INFO ";
        case TXLogLevel::WARN: return "WARN ";
        case TXLogLevel::ERROR: return "ERROR";
        case TXLogLevel::CRITICAL: return "CRIT ";
        default: return "UNKN ";
    }
}

// ==================== TXFileSink 实现 ====================

TXFileSink::TXFileSink(const std::string& filename, size_t max_size, int max_files)
    : filename_(filename), max_file_size_(max_size), current_size_(0), max_files_(max_files) {
    
    // 创建目录（如果不存在）
    std::filesystem::path file_path(filename);
    if (file_path.has_parent_path()) {
        std::filesystem::create_directories(file_path.parent_path());
    }
    
    file_.open(filename_, std::ios::app);
    if (file_.is_open()) {
        file_.seekp(0, std::ios::end);
        current_size_ = file_.tellp();
    }
}

TXFileSink::~TXFileSink() {
    if (file_.is_open()) {
        file_.close();
    }
}

void TXFileSink::write(const TXLogEntry& entry) {
    std::lock_guard<std::mutex> lock(write_mutex_);
    
    if (!file_.is_open()) return;
    
    std::string formatted = formatEntry(entry);
    file_ << formatted;
    current_size_ += formatted.size();
    
    // 检查是否需要轮转
    if (current_size_ > max_file_size_) {
        rotateFile();
    }
}

void TXFileSink::flush() {
    std::lock_guard<std::mutex> lock(write_mutex_);
    if (file_.is_open()) {
        file_.flush();
    }
}

void TXFileSink::rotateFile() {
    file_.close();
    
    // 轮转文件
    for (int i = max_files_ - 1; i > 0; --i) {
        std::string old_name = filename_ + "." + std::to_string(i);
        std::string new_name = filename_ + "." + std::to_string(i + 1);
        
        if (std::filesystem::exists(old_name)) {
            if (i == max_files_ - 1) {
                std::filesystem::remove(new_name); // 删除最老的文件
            }
            std::filesystem::rename(old_name, new_name);
        }
    }
    
    // 重命名当前文件
    std::string backup_name = filename_ + ".1";
    std::filesystem::rename(filename_, backup_name);
    
    // 创建新文件
    file_.open(filename_, std::ios::out);
    current_size_ = 0;
}

std::string TXFileSink::formatEntry(const TXLogEntry& entry) const {
    auto time_t = std::chrono::system_clock::to_time_t(entry.timestamp);
    auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(
        entry.timestamp.time_since_epoch()) % 1000;
    
    // 获取级别字符串
    std::string level_str;
    switch (entry.level) {
        case TXLogLevel::TRACE: level_str = "TRACE"; break;
        case TXLogLevel::DEBUG: level_str = "DEBUG"; break;
        case TXLogLevel::INFO: level_str = "INFO "; break;
        case TXLogLevel::WARN: level_str = "WARN "; break;
        case TXLogLevel::ERROR: level_str = "ERROR"; break;
        case TXLogLevel::CRITICAL: level_str = "CRIT "; break;
        default: level_str = "UNKN "; break;
    }

    return fmt::format("[{:%Y-%m-%d %H:%M:%S}.{:03d}] [{}] [{}] {}\n",
                      fmt::localtime(time_t), ms.count(),
                      level_str, entry.logger_name, entry.message);
}

// ==================== TXAsyncLogBuffer 实现 ====================

TXAsyncLogBuffer::TXAsyncLogBuffer(TXUnifiedMemoryManager& memory_manager)
    : memory_manager_(memory_manager), shutdown_(false) {
    worker_thread_ = std::thread(&TXAsyncLogBuffer::workerLoop, this);
}

TXAsyncLogBuffer::~TXAsyncLogBuffer() {
    shutdown();
}

void TXAsyncLogBuffer::addSink(std::shared_ptr<TXLogSink> sink) {
    sinks_.push_back(sink);
}

void TXAsyncLogBuffer::log(TXLogEntry entry) {
    if (shutdown_.load()) return;
    
    {
        std::lock_guard<std::mutex> lock(buffer_mutex_);
        buffer_.push(std::move(entry));
    }
    buffer_cv_.notify_one();
}

void TXAsyncLogBuffer::flush() {
    if (shutdown_.load()) return;

    // 🚀 使用超时等待避免死锁
    std::unique_lock<std::mutex> lock(buffer_mutex_);
    auto timeout = std::chrono::milliseconds(1000); // 1秒超时

    bool success = buffer_cv_.wait_for(lock, timeout, [this] {
        return buffer_.empty() || shutdown_.load();
    });

    if (!success) {
        std::cerr << "警告: 日志刷新超时" << std::endl;
    }

    // 刷新所有输出器
    for (auto& sink : sinks_) {
        try {
            sink->flush();
        } catch (const std::exception& e) {
            std::cerr << "日志输出器刷新失败: " << e.what() << std::endl;
        }
    }
}

void TXAsyncLogBuffer::shutdown() {
    if (shutdown_.exchange(true)) return;
    
    buffer_cv_.notify_all();
    if (worker_thread_.joinable()) {
        worker_thread_.join();
    }
}

void TXAsyncLogBuffer::workerLoop() {
    std::cout << "🚀 日志工作线程启动" << std::endl;

    while (!shutdown_.load()) {
        std::unique_lock<std::mutex> lock(buffer_mutex_);
        buffer_cv_.wait(lock, [this] { return !buffer_.empty() || shutdown_.load(); });

        if (shutdown_.load()) {
            std::cout << "🚀 日志工作线程收到关闭信号" << std::endl;
            break;
        }

        size_t processed = 0;
        while (!buffer_.empty()) {
            TXLogEntry entry = std::move(buffer_.front());
            buffer_.pop();
            lock.unlock();

            // 写入所有输出器
            for (auto& sink : sinks_) {
                try {
                    sink->write(entry);
                } catch (const std::exception& e) {
                    std::cerr << "日志写入失败: " << e.what() << std::endl;
                }
            }

            ++processed;
            lock.lock();
        }

        if (processed > 0) {
            // 通知等待的flush()
            buffer_cv_.notify_all();
        }
    }

    std::cout << "🚀 日志工作线程退出" << std::endl;
}

// ==================== TXHighPerformanceLogger 实现 ====================

TXHighPerformanceLogger::TXHighPerformanceLogger(const std::string& name, TXLogLevel level) {
    // 🚀 纯净版本：直接创建快速同步日志器，默认控制台输出
    fast_logger_ = std::make_unique<TXFastSyncLogger>(name, level, TXLogOutputMode::CONSOLE_ONLY, true);
}

void TXHighPerformanceLogger::flush() {
    if (fast_logger_) {
        fast_logger_->flush();
    }
}

// ==================== TXGlobalLogger 实现 ====================

std::shared_ptr<TXHighPerformanceLogger> TXGlobalLogger::default_logger_;
std::mutex TXGlobalLogger::logger_mutex_;

void TXGlobalLogger::initialize(TXUnifiedMemoryManager& memory_manager) {
    std::lock_guard<std::mutex> lock(logger_mutex_);

    if (!default_logger_) {
        // 🚀 纯净版本：直接创建高性能日志器，默认控制台输出
        default_logger_ = std::make_shared<TXHighPerformanceLogger>("TinaXlsx", TXLogLevel::INFO);

        std::cout << "🚀 TXGlobalLogger 初始化完成（纯净高性能版本）" << std::endl;
    }
}

void TXGlobalLogger::setOutputMode(TXLogOutputMode mode) {
    if (default_logger_ && default_logger_->fast_logger_) {
        default_logger_->fast_logger_->setOutputMode(mode);
        std::cout << "🚀 输出模式已设置为: " << static_cast<int>(mode) << std::endl;
    }
}

void TXGlobalLogger::shutdown() {
    std::lock_guard<std::mutex> lock(logger_mutex_);

    // 🚀 纯净版本：只需要重置日志器
    default_logger_.reset();
}

std::shared_ptr<TXHighPerformanceLogger> TXGlobalLogger::getDefault() {
    if (!default_logger_) {
        throw std::runtime_error("TXGlobalLogger not initialized! Call TXGlobalLogger::initialize() first.");
    }
    return default_logger_;
}

std::shared_ptr<TXHighPerformanceLogger> TXGlobalLogger::create(const std::string& name, TXLogLevel level) {
    // 🚀 纯净版本：直接创建独立的高性能日志器
    return std::make_shared<TXHighPerformanceLogger>(name, level);
}

} // namespace TinaXlsx
