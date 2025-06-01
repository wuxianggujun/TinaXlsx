#pragma once

#include <cstdint>
#include <vector>
#include <string>

namespace TinaXlsx {

/**
 * @brief 简单的SHA-512实现
 * 
 * 专门用于Excel密码哈希算法
 * 基于FIPS 180-4标准
 */
class TXSha512 {
public:
    TXSha512();
    ~TXSha512() = default;

    /**
     * @brief 重置哈希状态
     */
    void reset();

    /**
     * @brief 更新哈希数据
     * @param data 输入数据
     * @param length 数据长度
     */
    void update(const uint8_t* data, size_t length);

    /**
     * @brief 更新哈希数据
     * @param data 输入数据向量
     */
    void update(const std::vector<uint8_t>& data);

    /**
     * @brief 完成哈希计算并获取结果
     * @return 64字节的SHA-512哈希值
     */
    std::vector<uint8_t> finalize();

    /**
     * @brief 一次性计算SHA-512哈希
     * @param data 输入数据
     * @param length 数据长度
     * @return 64字节的SHA-512哈希值
     */
    static std::vector<uint8_t> hash(const uint8_t* data, size_t length);

    /**
     * @brief 一次性计算SHA-512哈希
     * @param data 输入数据向量
     * @return 64字节的SHA-512哈希值
     */
    static std::vector<uint8_t> hash(const std::vector<uint8_t>& data);

private:
    static constexpr size_t BLOCK_SIZE = 128;  // 1024 bits
    static constexpr size_t HASH_SIZE = 64;    // 512 bits

    uint64_t state_[8];
    uint8_t buffer_[BLOCK_SIZE];
    size_t bufferLength_;
    uint64_t totalLength_;

    void processBlock(const uint8_t* block);
    void pad();

    // SHA-512常量
    static const uint64_t K[80];
    
    // 辅助函数
    static uint64_t rotr(uint64_t x, int n);
    static uint64_t ch(uint64_t x, uint64_t y, uint64_t z);
    static uint64_t maj(uint64_t x, uint64_t y, uint64_t z);
    static uint64_t sigma0(uint64_t x);
    static uint64_t sigma1(uint64_t x);
    static uint64_t gamma0(uint64_t x);
    static uint64_t gamma1(uint64_t x);
};

/**
 * @brief Base64编码工具类
 */
class TXBase64 {
public:
    /**
     * @brief 将字节数组编码为Base64字符串
     * @param data 输入数据
     * @return Base64编码的字符串
     */
    static std::string encode(const std::vector<uint8_t>& data);

    /**
     * @brief 将Base64字符串解码为字节数组
     * @param encoded Base64编码的字符串
     * @return 解码后的字节数组
     */
    static std::vector<uint8_t> decode(const std::string& encoded);

private:
    static const std::string CHARS;
    static const int DECODE_TABLE[128];
};

/**
 * @brief Excel密码哈希工具类
 */
class TXExcelPasswordHash {
public:
    /**
     * @brief 生成随机盐值
     * @param length 盐值长度（字节）
     * @return Base64编码的盐值
     */
    static std::string generateSalt(size_t length = 16);

    /**
     * @brief 将密码转换为UTF-16编码（小端序，无BOM）
     * @param password UTF-8密码字符串
     * @return UTF-16编码的字节数组
     */
    static std::vector<uint8_t> passwordToUtf16(const std::string& password);

    /**
     * @brief 计算Excel密码哈希
     * @param password 密码字符串
     * @param saltBase64 Base64编码的盐值
     * @param spinCount 迭代次数（默认100000）
     * @return Base64编码的哈希值
     */
    static std::string calculateHash(const std::string& password, 
                                   const std::string& saltBase64, 
                                   uint32_t spinCount = 100000);

    /**
     * @brief 验证密码
     * @param password 输入的密码
     * @param saltBase64 Base64编码的盐值
     * @param hashBase64 Base64编码的哈希值
     * @param spinCount 迭代次数
     * @return 验证成功返回true
     */
    static bool verifyPassword(const std::string& password,
                             const std::string& saltBase64,
                             const std::string& hashBase64,
                             uint32_t spinCount = 100000);
};

} // namespace TinaXlsx
