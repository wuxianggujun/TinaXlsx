#include "TinaXlsx/TXSha512.hpp"
#include <cstring>
#include <random>
#include <sstream>
#include <iomanip>
#include <codecvt>
#include <locale>

namespace TinaXlsx {

// SHA-512常量K
const uint64_t TXSha512::K[80] = {
    0x428a2f98d728ae22ULL, 0x7137449123ef65cdULL, 0xb5c0fbcfec4d3b2fULL, 0xe9b5dba58189dbbcULL,
    0x3956c25bf348b538ULL, 0x59f111f1b605d019ULL, 0x923f82a4af194f9bULL, 0xab1c5ed5da6d8118ULL,
    0xd807aa98a3030242ULL, 0x12835b0145706fbeULL, 0x243185be4ee4b28cULL, 0x550c7dc3d5ffb4e2ULL,
    0x72be5d74f27b896fULL, 0x80deb1fe3b1696b1ULL, 0x9bdc06a725c71235ULL, 0xc19bf174cf692694ULL,
    0xe49b69c19ef14ad2ULL, 0xefbe4786384f25e3ULL, 0x0fc19dc68b8cd5b5ULL, 0x240ca1cc77ac9c65ULL,
    0x2de92c6f592b0275ULL, 0x4a7484aa6ea6e483ULL, 0x5cb0a9dcbd41fbd4ULL, 0x76f988da831153b5ULL,
    0x983e5152ee66dfabULL, 0xa831c66d2db43210ULL, 0xb00327c898fb213fULL, 0xbf597fc7beef0ee4ULL,
    0xc6e00bf33da88fc2ULL, 0xd5a79147930aa725ULL, 0x06ca6351e003826fULL, 0x142929670a0e6e70ULL,
    0x27b70a8546d22ffcULL, 0x2e1b21385c26c926ULL, 0x4d2c6dfc5ac42aedULL, 0x53380d139d95b3dfULL,
    0x650a73548baf63deULL, 0x766a0abb3c77b2a8ULL, 0x81c2c92e47edaee6ULL, 0x92722c851482353bULL,
    0xa2bfe8a14cf10364ULL, 0xa81a664bbc423001ULL, 0xc24b8b70d0f89791ULL, 0xc76c51a30654be30ULL,
    0xd192e819d6ef5218ULL, 0xd69906245565a910ULL, 0xf40e35855771202aULL, 0x106aa07032bbd1b8ULL,
    0x19a4c116b8d2d0c8ULL, 0x1e376c085141ab53ULL, 0x2748774cdf8eeb99ULL, 0x34b0bcb5e19b48a8ULL,
    0x391c0cb3c5c95a63ULL, 0x4ed8aa4ae3418acbULL, 0x5b9cca4f7763e373ULL, 0x682e6ff3d6b2b8a3ULL,
    0x748f82ee5defb2fcULL, 0x78a5636f43172f60ULL, 0x84c87814a1f0ab72ULL, 0x8cc702081a6439ecULL,
    0x90befffa23631e28ULL, 0xa4506cebde82bde9ULL, 0xbef9a3f7b2c67915ULL, 0xc67178f2e372532bULL,
    0xca273eceea26619cULL, 0xd186b8c721c0c207ULL, 0xeada7dd6cde0eb1eULL, 0xf57d4f7fee6ed178ULL,
    0x06f067aa72176fbaULL, 0x0a637dc5a2c898a6ULL, 0x113f9804bef90daeULL, 0x1b710b35131c471bULL,
    0x28db77f523047d84ULL, 0x32caab7b40c72493ULL, 0x3c9ebe0a15c9bebcULL, 0x431d67c49c100d4cULL,
    0x4cc5d4becb3e42b6ULL, 0x597f299cfc657e2aULL, 0x5fcb6fab3ad6faecULL, 0x6c44198c4a475817ULL
};

TXSha512::TXSha512() {
    reset();
}

void TXSha512::reset() {
    // SHA-512初始哈希值
    state_[0] = 0x6a09e667f3bcc908ULL;
    state_[1] = 0xbb67ae8584caa73bULL;
    state_[2] = 0x3c6ef372fe94f82bULL;
    state_[3] = 0xa54ff53a5f1d36f1ULL;
    state_[4] = 0x510e527fade682d1ULL;
    state_[5] = 0x9b05688c2b3e6c1fULL;
    state_[6] = 0x1f83d9abfb41bd6bULL;
    state_[7] = 0x5be0cd19137e2179ULL;
    
    bufferLength_ = 0;
    totalLength_ = 0;
}

uint64_t TXSha512::rotr(uint64_t x, int n) {
    return (x >> n) | (x << (64 - n));
}

uint64_t TXSha512::ch(uint64_t x, uint64_t y, uint64_t z) {
    return (x & y) ^ (~x & z);
}

uint64_t TXSha512::maj(uint64_t x, uint64_t y, uint64_t z) {
    return (x & y) ^ (x & z) ^ (y & z);
}

uint64_t TXSha512::sigma0(uint64_t x) {
    return rotr(x, 28) ^ rotr(x, 34) ^ rotr(x, 39);
}

uint64_t TXSha512::sigma1(uint64_t x) {
    return rotr(x, 14) ^ rotr(x, 18) ^ rotr(x, 41);
}

uint64_t TXSha512::gamma0(uint64_t x) {
    return rotr(x, 1) ^ rotr(x, 8) ^ (x >> 7);
}

uint64_t TXSha512::gamma1(uint64_t x) {
    return rotr(x, 19) ^ rotr(x, 61) ^ (x >> 6);
}

void TXSha512::update(const uint8_t* data, size_t length) {
    totalLength_ += length;
    
    while (length > 0) {
        size_t copyLength = std::min(length, BLOCK_SIZE - bufferLength_);
        std::memcpy(buffer_ + bufferLength_, data, copyLength);
        bufferLength_ += copyLength;
        data += copyLength;
        length -= copyLength;
        
        if (bufferLength_ == BLOCK_SIZE) {
            processBlock(buffer_);
            bufferLength_ = 0;
        }
    }
}

void TXSha512::update(const std::vector<uint8_t>& data) {
    update(data.data(), data.size());
}

void TXSha512::processBlock(const uint8_t* block) {
    uint64_t W[80];
    
    // 准备消息调度
    for (int i = 0; i < 16; i++) {
        W[i] = ((uint64_t)block[i * 8] << 56) |
               ((uint64_t)block[i * 8 + 1] << 48) |
               ((uint64_t)block[i * 8 + 2] << 40) |
               ((uint64_t)block[i * 8 + 3] << 32) |
               ((uint64_t)block[i * 8 + 4] << 24) |
               ((uint64_t)block[i * 8 + 5] << 16) |
               ((uint64_t)block[i * 8 + 6] << 8) |
               ((uint64_t)block[i * 8 + 7]);
    }
    
    for (int i = 16; i < 80; i++) {
        W[i] = gamma1(W[i - 2]) + W[i - 7] + gamma0(W[i - 15]) + W[i - 16];
    }
    
    // 初始化工作变量
    uint64_t a = state_[0], b = state_[1], c = state_[2], d = state_[3];
    uint64_t e = state_[4], f = state_[5], g = state_[6], h = state_[7];
    
    // 主循环
    for (int i = 0; i < 80; i++) {
        uint64_t T1 = h + sigma1(e) + ch(e, f, g) + K[i] + W[i];
        uint64_t T2 = sigma0(a) + maj(a, b, c);
        h = g;
        g = f;
        f = e;
        e = d + T1;
        d = c;
        c = b;
        b = a;
        a = T1 + T2;
    }
    
    // 更新哈希值
    state_[0] += a;
    state_[1] += b;
    state_[2] += c;
    state_[3] += d;
    state_[4] += e;
    state_[5] += f;
    state_[6] += g;
    state_[7] += h;
}

void TXSha512::pad() {
    uint64_t totalBits = totalLength_ * 8;
    
    // 添加填充位
    buffer_[bufferLength_++] = 0x80;
    
    // 如果剩余空间不足16字节，需要额外的块
    if (bufferLength_ > BLOCK_SIZE - 16) {
        while (bufferLength_ < BLOCK_SIZE) {
            buffer_[bufferLength_++] = 0;
        }
        processBlock(buffer_);
        bufferLength_ = 0;
    }
    
    // 填充零直到剩余16字节
    while (bufferLength_ < BLOCK_SIZE - 16) {
        buffer_[bufferLength_++] = 0;
    }
    
    // 添加长度（128位，高64位为0）
    for (int i = 0; i < 8; i++) {
        buffer_[bufferLength_++] = 0;
    }
    for (int i = 7; i >= 0; i--) {
        buffer_[bufferLength_++] = (totalBits >> (i * 8)) & 0xFF;
    }
}

std::vector<uint8_t> TXSha512::finalize() {
    pad();
    processBlock(buffer_);
    
    std::vector<uint8_t> result(HASH_SIZE);
    for (int i = 0; i < 8; i++) {
        for (int j = 7; j >= 0; j--) {
            result[i * 8 + (7 - j)] = (state_[i] >> (j * 8)) & 0xFF;
        }
    }
    
    reset();
    return result;
}

std::vector<uint8_t> TXSha512::hash(const uint8_t* data, size_t length) {
    TXSha512 sha;
    sha.update(data, length);
    return sha.finalize();
}

std::vector<uint8_t> TXSha512::hash(const std::vector<uint8_t>& data) {
    return hash(data.data(), data.size());
}

// ==================== Base64编码实现 ====================

const std::string TXBase64::CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

const int TXBase64::DECODE_TABLE[128] = {
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,62,-1,-1,-1,63,
    52,53,54,55,56,57,58,59,60,61,-1,-1,-1,-1,-1,-1,
    -1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9,10,11,12,13,14,
    15,16,17,18,19,20,21,22,23,24,25,-1,-1,-1,-1,-1,
    -1,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,
    41,42,43,44,45,46,47,48,49,50,51,-1,-1,-1,-1,-1
};

std::string TXBase64::encode(const std::vector<uint8_t>& data) {
    std::string result;
    size_t i = 0;

    while (i < data.size()) {
        uint32_t octet_a = i < data.size() ? data[i++] : 0;
        uint32_t octet_b = i < data.size() ? data[i++] : 0;
        uint32_t octet_c = i < data.size() ? data[i++] : 0;

        uint32_t triple = (octet_a << 0x10) + (octet_b << 0x08) + octet_c;

        result += CHARS[(triple >> 3 * 6) & 0x3F];
        result += CHARS[(triple >> 2 * 6) & 0x3F];
        result += CHARS[(triple >> 1 * 6) & 0x3F];
        result += CHARS[(triple >> 0 * 6) & 0x3F];
    }

    // 添加填充
    size_t padding = (3 - data.size() % 3) % 3;
    for (size_t j = 0; j < padding; j++) {
        result[result.length() - 1 - j] = '=';
    }

    return result;
}

std::vector<uint8_t> TXBase64::decode(const std::string& encoded) {
    std::vector<uint8_t> result;

    for (size_t i = 0; i < encoded.length(); i += 4) {
        uint32_t sextet_a = (i < encoded.length() && encoded[i] != '=') ? DECODE_TABLE[static_cast<uint8_t>(encoded[i])] : 0;
        uint32_t sextet_b = (i + 1 < encoded.length() && encoded[i + 1] != '=') ? DECODE_TABLE[static_cast<uint8_t>(encoded[i + 1])] : 0;
        uint32_t sextet_c = (i + 2 < encoded.length() && encoded[i + 2] != '=') ? DECODE_TABLE[static_cast<uint8_t>(encoded[i + 2])] : 0;
        uint32_t sextet_d = (i + 3 < encoded.length() && encoded[i + 3] != '=') ? DECODE_TABLE[static_cast<uint8_t>(encoded[i + 3])] : 0;

        uint32_t triple = (sextet_a << 18) + (sextet_b << 12) + (sextet_c << 6) + sextet_d;

        if (i < encoded.length()) {
            result.push_back((triple >> 16) & 0xFF);
        }
        if (i + 1 < encoded.length() && encoded[i + 2] != '=') {
            result.push_back((triple >> 8) & 0xFF);
        }
        if (i + 2 < encoded.length() && encoded[i + 3] != '=') {
            result.push_back(triple & 0xFF);
        }
    }

    return result;
}

// ==================== Excel密码哈希实现 ====================

std::string TXExcelPasswordHash::generateSalt(size_t length) {
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_int_distribution<> dis(0, 255);

    std::vector<uint8_t> salt(length);
    for (size_t i = 0; i < length; i++) {
        salt[i] = static_cast<uint8_t>(dis(gen));
    }

    return TXBase64::encode(salt);
}

std::vector<uint8_t> TXExcelPasswordHash::passwordToUtf16(const std::string& password) {
    std::vector<uint8_t> result;

    // 简单的ASCII到UTF-16LE转换（对于基本ASCII字符）
    for (char c : password) {
        result.push_back(static_cast<uint8_t>(c));  // 低字节
        result.push_back(0);                        // 高字节
    }

    return result;
}

std::string TXExcelPasswordHash::calculateHash(const std::string& password,
                                             const std::string& saltBase64,
                                             uint32_t spinCount) {
    // 解码盐值
    std::vector<uint8_t> salt = TXBase64::decode(saltBase64);

    // 将密码转换为UTF-16LE
    std::vector<uint8_t> passwordBytes = passwordToUtf16(password);

    // 第一步：salt + password
    std::vector<uint8_t> buffer;
    buffer.insert(buffer.end(), salt.begin(), salt.end());
    buffer.insert(buffer.end(), passwordBytes.begin(), passwordBytes.end());

    // 计算初始哈希
    std::vector<uint8_t> hash = TXSha512::hash(buffer);

    // 迭代spinCount次
    for (uint32_t i = 0; i < spinCount; i++) {
        buffer.clear();
        buffer.insert(buffer.end(), hash.begin(), hash.end());

        // 添加迭代计数器（小端序）
        buffer.push_back(i & 0xFF);
        buffer.push_back((i >> 8) & 0xFF);
        buffer.push_back((i >> 16) & 0xFF);
        buffer.push_back((i >> 24) & 0xFF);

        hash = TXSha512::hash(buffer);
    }

    return TXBase64::encode(hash);
}

bool TXExcelPasswordHash::verifyPassword(const std::string& password,
                                       const std::string& saltBase64,
                                       const std::string& hashBase64,
                                       uint32_t spinCount) {
    std::string calculatedHash = calculateHash(password, saltBase64, spinCount);
    return calculatedHash == hashBase64;
}

} // namespace TinaXlsx
