#include "TinaXlsx/TinaXlsx.hpp"
#include <mutex>
#include <atomic>

namespace TinaXlsx {

namespace {
    std::atomic<bool> g_initialized{false};
    std::mutex g_init_mutex;
}

bool initialize() {
    std::lock_guard<std::mutex> lock(g_init_mutex);
    
    if (g_initialized.load()) {
        return true;
    }
    
    // 这里可以添加任何必要的初始化代码
    // 比如设置默认配置、初始化线程池等
    
    g_initialized.store(true);
    return true;
}

void cleanup() {
    std::lock_guard<std::mutex> lock(g_init_mutex);
    
    if (!g_initialized.load()) {
        return;
    }
    
    // 这里可以添加任何必要的清理代码
    // 比如清理全局资源、停止线程池等
    
    g_initialized.store(false);
}

std::string getVersion() {
    return Version::getString();
}

std::string getBuildInfo() {
    return "TinaXlsx " + Version::getString() + 
           " built on " + __DATE__ + " " + __TIME__ +
           " with C++17";
}

} // namespace TinaXlsx 