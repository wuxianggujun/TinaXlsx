//
// @file TXHighPerformanceLogger.hpp
// @brief 🚀 高性能日志库 - 基于fmt库的零分配日志系统
//

#pragma once

#include "TXUnifiedMemoryManager.hpp"
#include <fmt/format.h>
#include <fmt/chrono.h>
#include <string>
#include <fstream>
#include <iostream>
#include <memory>
#include <atomic>
#include <thread>
#include <queue>
#include <mutex>
#include <condition_variable>
#include <chrono>

namespace TinaXlsx {

/**
 * @brief 日志级别
 */
enum class TXLogLevel : uint8_t {
    TRACE = 0,
    DEBUG = 1,
    INFO = 2,
    WARN = 3,
    ERROR = 4,
    CRITICAL = 5,
    OFF = 6
};

/**
 * @brief 🚀 日志输出模式
 */
enum class TXLogOutputMode : uint8_t {
    CONSOLE_ONLY = 0,    // 仅控制台（最快）
    FILE_ONLY = 1,       // 仅文件
    BOTH = 2,           // 控制台+文件（最慢但最完整）
    PERFORMANCE = 3      // 性能模式（基准测试时使用）
};

/**
 * @brief 🚀 高性能日志条目
 */
struct TXLogEntry {
    TXLogLevel level;
    std::chrono::system_clock::time_point timestamp;
    std::string_view logger_name;
    std::string message;
    std::string_view file;
    int line;
    std::string_view function;
    
    TXLogEntry() = default;
    TXLogEntry(TXLogLevel lvl, std::string_view name, std::string msg, 
               std::string_view f, int l, std::string_view func)
        : level(lvl), timestamp(std::chrono::system_clock::now()), 
          logger_name(name), message(std::move(msg)), file(f), line(l), function(func) {}
};

/**
 * @brief 🚀 日志输出器接口
 */
class TXLogSink {
public:
    virtual ~TXLogSink() = default;
    virtual void write(const TXLogEntry& entry) = 0;
    virtual void flush() = 0;
};

/**
 * @brief 🚀 控制台输出器
 */
class TXConsoleSink : public TXLogSink {
private:
    bool colored_output_;
    std::mutex write_mutex_;
    
public:
    explicit TXConsoleSink(bool colored = true);
    void write(const TXLogEntry& entry) override;
    void flush() override;
    
private:
    std::string getColorCode(TXLogLevel level) const;
    std::string getLevelString(TXLogLevel level) const;
};

/**
 * @brief 🚀 高性能文件输出器
 */
class TXFileSink : public TXLogSink {
private:
    std::string filename_;
    std::ofstream file_;
    std::mutex write_mutex_;
    size_t max_file_size_;
    size_t current_size_;
    int max_files_;
    
public:
    TXFileSink(const std::string& filename, size_t max_size = 10 * 1024 * 1024, int max_files = 5);
    ~TXFileSink();
    
    void write(const TXLogEntry& entry) override;
    void flush() override;
    
private:
    void rotateFile();
    std::string formatEntry(const TXLogEntry& entry) const;
};

/**
 * @brief 🚀 异步日志缓冲区
 */
class TXAsyncLogBuffer {
private:
    TXUnifiedMemoryManager& memory_manager_;
    std::queue<TXLogEntry> buffer_;
    std::mutex buffer_mutex_;
    std::condition_variable buffer_cv_;
    std::atomic<bool> shutdown_;
    std::thread worker_thread_;
    std::vector<std::shared_ptr<TXLogSink>> sinks_;
    
public:
    explicit TXAsyncLogBuffer(TXUnifiedMemoryManager& memory_manager);
    ~TXAsyncLogBuffer();
    
    void addSink(std::shared_ptr<TXLogSink> sink);
    void log(TXLogEntry entry);
    void flush();
    void shutdown();
    
private:
    void workerLoop();
};

/**
 * @brief 🚀 超高性能同步日志器
 */
class TXFastSyncLogger {
private:
    std::string name_;
    TXLogLevel level_;
    TXLogOutputMode output_mode_;
    bool colored_output_;
    std::unique_ptr<std::ofstream> file_stream_;

public:
    TXFastSyncLogger(const std::string& name, TXLogLevel level = TXLogLevel::INFO,
                     TXLogOutputMode mode = TXLogOutputMode::CONSOLE_ONLY, bool colored = true)
        : name_(name), level_(level), output_mode_(mode), colored_output_(colored) {

        // 如果需要文件输出，创建文件流
        if (output_mode_ == TXLogOutputMode::FILE_ONLY || output_mode_ == TXLogOutputMode::BOTH) {
            file_stream_ = std::make_unique<std::ofstream>("tinaxlsx.log", std::ios::app);
        }
    }

    void setLevel(TXLogLevel level) { level_ = level; }
    TXLogLevel getLevel() const { return level_; }

    // 🚀 设置输出模式
    void setOutputMode(TXLogOutputMode mode) {
        output_mode_ = mode;
        if ((mode == TXLogOutputMode::FILE_ONLY || mode == TXLogOutputMode::BOTH) && !file_stream_) {
            file_stream_ = std::make_unique<std::ofstream>("tinaxlsx.log", std::ios::app);
        }
    }

    void flush() {
        // 🚀 纯净的刷新：只处理我们自己的输出
        if (output_mode_ == TXLogOutputMode::CONSOLE_ONLY || output_mode_ == TXLogOutputMode::BOTH ||
            output_mode_ == TXLogOutputMode::PERFORMANCE) {
            std::cout.flush();
        }
        if ((output_mode_ == TXLogOutputMode::FILE_ONLY || output_mode_ == TXLogOutputMode::BOTH) && file_stream_) {
            file_stream_->flush();
        }
    }

    // 🚀 超高性能模板化日志方法 - 直接输出，无异步开销
    template<typename... Args>
    void trace(fmt::format_string<Args...> fmt, Args&&... args) {
        if (TXLogLevel::TRACE >= level_) {
            logDirect(TXLogLevel::TRACE, fmt, std::forward<Args>(args)...);
        }
    }

    template<typename... Args>
    void debug(fmt::format_string<Args...> fmt, Args&&... args) {
        if (TXLogLevel::DEBUG >= level_) {
            logDirect(TXLogLevel::DEBUG, fmt, std::forward<Args>(args)...);
        }
    }

    template<typename... Args>
    void info(fmt::format_string<Args...> fmt, Args&&... args) {
        if (TXLogLevel::INFO >= level_) {
            logDirect(TXLogLevel::INFO, fmt, std::forward<Args>(args)...);
        }
    }

    template<typename... Args>
    void warn(fmt::format_string<Args...> fmt, Args&&... args) {
        if (TXLogLevel::WARN >= level_) {
            logDirect(TXLogLevel::WARN, fmt, std::forward<Args>(args)...);
        }
    }

    template<typename... Args>
    void error(fmt::format_string<Args...> fmt, Args&&... args) {
        if (TXLogLevel::ERROR >= level_) {
            logDirect(TXLogLevel::ERROR, fmt, std::forward<Args>(args)...);
        }
    }

    template<typename... Args>
    void critical(fmt::format_string<Args...> fmt, Args&&... args) {
        if (TXLogLevel::CRITICAL >= level_) {
            logDirect(TXLogLevel::CRITICAL, fmt, std::forward<Args>(args)...);
        }
    }

private:
    template<typename... Args>
    void logDirect(TXLogLevel level, fmt::format_string<Args...> fmt, Args&&... args) {
        // 🚀 超高性能优化：最小化所有开销

        // 🚀 性能模式：极简输出，最接近std::cout
        if (output_mode_ == TXLogOutputMode::PERFORMANCE) {
            // 🚀 最简实现：直接输出，最小化开销
            std::cout << '[' << getLevelChar(level) << "] ";
            fmt::print(fmt, std::forward<Args>(args)...);
            std::cout << '\n';
            return;
        }

        // 🚀 控制台输出（优化版）
        if (output_mode_ == TXLogOutputMode::CONSOLE_ONLY || output_mode_ == TXLogOutputMode::BOTH) {
            if (colored_output_) {
                std::cout << getColorCode(level);
            }

            // 🚀 减少字符串操作，直接输出
            std::cout << '[' << getLevelChar(level) << "] [" << name_ << "] ";
            fmt::print(fmt, std::forward<Args>(args)...);
            std::cout << '\n';

            if (colored_output_) {
                std::cout << "\033[0m"; // 重置颜色
            }
        }

        // 🚀 文件输出（仅在需要时）
        if ((output_mode_ == TXLogOutputMode::FILE_ONLY || output_mode_ == TXLogOutputMode::BOTH) && file_stream_) {
            // 🚀 简化时间戳，减少开销
            auto now = std::chrono::system_clock::now();
            auto time_t = std::chrono::system_clock::to_time_t(now);

            // 🚀 使用传统的流操作，避免fmt::print的文件流问题
            *file_stream_ << '[' << getLevelChar(level) << "] [" << name_ << "] ";
            *file_stream_ << fmt::format(fmt, std::forward<Args>(args)...);
            *file_stream_ << '\n';
        }
    }

    // 🚀 超快级别字符获取（单字符，最快）
    char getLevelChar(TXLogLevel level) const {
        switch (level) {
            case TXLogLevel::TRACE: return 'T';
            case TXLogLevel::DEBUG: return 'D';
            case TXLogLevel::INFO: return 'I';
            case TXLogLevel::WARN: return 'W';
            case TXLogLevel::ERROR: return 'E';
            case TXLogLevel::CRITICAL: return 'C';
            default: return 'U';
        }
    }

    const char* getLevelString(TXLogLevel level) const {
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

    const char* getColorCode(TXLogLevel level) const {
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
};

/**
 * @brief 🚀 高性能日志器（保持兼容性）
 */
class TXHighPerformanceLogger {
    friend class TXGlobalLogger; // 🚀 允许TXGlobalLogger访问私有成员

private:
    std::unique_ptr<TXFastSyncLogger> fast_logger_; // 🚀 纯净版本：只需要快速日志器

public:
    TXHighPerformanceLogger(const std::string& name, TXLogLevel level = TXLogLevel::INFO);

    void setLevel(TXLogLevel level) {
        if (fast_logger_) fast_logger_->setLevel(level);
    }
    TXLogLevel getLevel() const {
        return fast_logger_ ? fast_logger_->getLevel() : TXLogLevel::INFO;
    }

    void flush();

    // 🚀 设置输出模式
    void setOutputMode(TXLogOutputMode mode) {
        if (fast_logger_) fast_logger_->setOutputMode(mode);
    }

    // 🚀 高性能模板化日志方法
    template<typename... Args>
    void trace(fmt::format_string<Args...> fmt, Args&&... args) {
        log(TXLogLevel::TRACE, fmt, std::forward<Args>(args)...);
    }

    template<typename... Args>
    void debug(fmt::format_string<Args...> fmt, Args&&... args) {
        log(TXLogLevel::DEBUG, fmt, std::forward<Args>(args)...);
    }

    template<typename... Args>
    void info(fmt::format_string<Args...> fmt, Args&&... args) {
        log(TXLogLevel::INFO, fmt, std::forward<Args>(args)...);
    }

    template<typename... Args>
    void warn(fmt::format_string<Args...> fmt, Args&&... args) {
        log(TXLogLevel::WARN, fmt, std::forward<Args>(args)...);
    }

    template<typename... Args>
    void error(fmt::format_string<Args...> fmt, Args&&... args) {
        log(TXLogLevel::ERROR, fmt, std::forward<Args>(args)...);
    }

    template<typename... Args>
    void critical(fmt::format_string<Args...> fmt, Args&&... args) {
        log(TXLogLevel::CRITICAL, fmt, std::forward<Args>(args)...);
    }

private:
    template<typename... Args>
    void log(TXLogLevel level, fmt::format_string<Args...> fmt, Args&&... args) {
        // 🚀 纯净版本：直接委托给快速日志器
        if (!fast_logger_) return;

        switch (level) {
            case TXLogLevel::TRACE: fast_logger_->trace(fmt, std::forward<Args>(args)...); break;
            case TXLogLevel::DEBUG: fast_logger_->debug(fmt, std::forward<Args>(args)...); break;
            case TXLogLevel::INFO: fast_logger_->info(fmt, std::forward<Args>(args)...); break;
            case TXLogLevel::WARN: fast_logger_->warn(fmt, std::forward<Args>(args)...); break;
            case TXLogLevel::ERROR: fast_logger_->error(fmt, std::forward<Args>(args)...); break;
            case TXLogLevel::CRITICAL: fast_logger_->critical(fmt, std::forward<Args>(args)...); break;
            default: break;
        }
    }
};

/**
 * @brief 🚀 全局日志管理器
 */
class TXGlobalLogger {
private:
    static std::shared_ptr<TXHighPerformanceLogger> default_logger_;
    static std::mutex logger_mutex_;
    
public:
    static void initialize(TXUnifiedMemoryManager& memory_manager);
    static void shutdown();
    
    static std::shared_ptr<TXHighPerformanceLogger> getDefault();
    static std::shared_ptr<TXHighPerformanceLogger> create(const std::string& name,
                                                           TXLogLevel level = TXLogLevel::INFO);

    // 🚀 设置全局输出模式
    static void setOutputMode(TXLogOutputMode mode);
    
    // 🚀 便捷的全局日志宏
    template<typename... Args>
    static void trace(fmt::format_string<Args...> fmt, Args&&... args) {
        getDefault()->trace(fmt, std::forward<Args>(args)...);
    }
    
    template<typename... Args>
    static void debug(fmt::format_string<Args...> fmt, Args&&... args) {
        getDefault()->debug(fmt, std::forward<Args>(args)...);
    }
    
    template<typename... Args>
    static void info(fmt::format_string<Args...> fmt, Args&&... args) {
        getDefault()->info(fmt, std::forward<Args>(args)...);
    }
    
    template<typename... Args>
    static void warn(fmt::format_string<Args...> fmt, Args&&... args) {
        getDefault()->warn(fmt, std::forward<Args>(args)...);
    }
    
    template<typename... Args>
    static void error(fmt::format_string<Args...> fmt, Args&&... args) {
        getDefault()->error(fmt, std::forward<Args>(args)...);
    }
    
    template<typename... Args>
    static void critical(fmt::format_string<Args...> fmt, Args&&... args) {
        getDefault()->critical(fmt, std::forward<Args>(args)...);
    }
};

} // namespace TinaXlsx

// 🚀 便捷的日志宏
#define TX_LOG_TRACE(...) TinaXlsx::TXGlobalLogger::trace(__VA_ARGS__)
#define TX_LOG_DEBUG(...) TinaXlsx::TXGlobalLogger::debug(__VA_ARGS__)
#define TX_LOG_INFO(...) TinaXlsx::TXGlobalLogger::info(__VA_ARGS__)
#define TX_LOG_WARN(...) TinaXlsx::TXGlobalLogger::warn(__VA_ARGS__)
#define TX_LOG_ERROR(...) TinaXlsx::TXGlobalLogger::error(__VA_ARGS__)
#define TX_LOG_CRITICAL(...) TinaXlsx::TXGlobalLogger::critical(__VA_ARGS__)

// 🚀 性能测试专用宏
#define TX_PERF_LOG(level, ...) \
    do { \
        auto start = std::chrono::high_resolution_clock::now(); \
        __VA_ARGS__; \
        auto end = std::chrono::high_resolution_clock::now(); \
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start).count() / 1000.0; \
        TinaXlsx::TXGlobalLogger::level("⏱️ Performance: {} took {:.3f}ms", #__VA_ARGS__, duration); \
    } while(0)
