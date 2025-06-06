//
// @file TXVector.hpp
// @brief 🚀 高性能向量容器 - 使用TXUnifiedMemoryManager的极致性能vector
//

#pragma once

#include "TXUnifiedMemoryManager.hpp"
#include <cstring>
#include <stdexcept>
#include <initializer_list>
#include <algorithm>
#include <iostream>

namespace TinaXlsx {

/**
 * @brief 🚀 高性能向量容器 - 专为TXUnifiedMemoryManager优化
 * 
 * 特性：
 * - 使用TXUnifiedMemoryManager进行内存分配
 * - SIMD友好的内存对齐
 * - 批量操作优化
 * - 零开销抽象
 * - 与std::vector兼容的接口
 */
template<typename T>
class TXVector {
public:
    using value_type = T;
    using size_type = size_t;
    using difference_type = ptrdiff_t;
    using reference = T&;
    using const_reference = const T&;
    using pointer = T*;
    using const_pointer = const T*;
    using iterator = T*;
    using const_iterator = const T*;

private:
    TXUnifiedMemoryManager* memory_manager_;
    T* data_;
    size_type size_;
    size_type capacity_;

    // 🚀 SIMD对齐常量
    static constexpr size_type SIMD_ALIGNMENT = 32; // AVX2对齐
    static constexpr size_type MIN_CAPACITY = 16;

public:
    // ==================== 构造函数 ====================
    
    /**
     * @brief 默认构造函数
     */
    explicit TXVector(TXUnifiedMemoryManager& manager) 
        : memory_manager_(&manager), data_(nullptr), size_(0), capacity_(0) {}
    
    /**
     * @brief 带初始容量的构造函数
     */
    TXVector(TXUnifiedMemoryManager& manager, size_type initial_capacity)
        : memory_manager_(&manager), data_(nullptr), size_(0), capacity_(0) {
        reserve(initial_capacity);
    }
    
    /**
     * @brief 带初始大小和值的构造函数
     */
    TXVector(TXUnifiedMemoryManager& manager, size_type count, const T& value)
        : memory_manager_(&manager), data_(nullptr), size_(0), capacity_(0) {
        assign(count, value);
    }
    
    /**
     * @brief 初始化列表构造函数
     */
    TXVector(TXUnifiedMemoryManager& manager, std::initializer_list<T> init)
        : memory_manager_(&manager), data_(nullptr), size_(0), capacity_(0) {
        assign(init);
    }
    
    /**
     * @brief 拷贝构造函数
     */
    TXVector(const TXVector& other)
        : memory_manager_(other.memory_manager_), data_(nullptr), size_(0), capacity_(0) {
        assign(other.begin(), other.end());
    }
    
    /**
     * @brief 移动构造函数
     */
    TXVector(TXVector&& other) noexcept
        : memory_manager_(other.memory_manager_), data_(other.data_), 
          size_(other.size_), capacity_(other.capacity_) {
        other.data_ = nullptr;
        other.size_ = 0;
        other.capacity_ = 0;
    }
    
    /**
     * @brief 析构函数
     */
    ~TXVector() {
        clear();
        deallocate();
    }
    
    // ==================== 赋值操作符 ====================
    
    TXVector& operator=(const TXVector& other) {
        if (this != &other) {
            assign(other.begin(), other.end());
        }
        return *this;
    }
    
    TXVector& operator=(TXVector&& other) noexcept {
        if (this != &other) {
            clear();
            deallocate();
            
            memory_manager_ = other.memory_manager_;
            data_ = other.data_;
            size_ = other.size_;
            capacity_ = other.capacity_;
            
            other.data_ = nullptr;
            other.size_ = 0;
            other.capacity_ = 0;
        }
        return *this;
    }
    
    TXVector& operator=(std::initializer_list<T> init) {
        assign(init);
        return *this;
    }
    
    // ==================== 容量管理 ====================
    
    size_type size() const noexcept { return size_; }
    size_type capacity() const noexcept { return capacity_; }
    bool empty() const noexcept { return size_ == 0; }
    
    /**
     * @brief 🚀 高性能预分配 - 使用智能增长策略
     */
    void reserve(size_type new_capacity) {
        if (new_capacity <= capacity_) return;
        
        // 🚀 智能容量增长：SIMD对齐 + 增长因子
        size_type aligned_capacity = alignCapacity(new_capacity);
        reallocate(aligned_capacity);
    }
    
    /**
     * @brief 🚀 高性能调整大小
     */
    void resize(size_type new_size) {
        if (new_size > capacity_) {
            reserve(new_size);
        }
        
        if (new_size > size_) {
            // 默认构造新元素
            for (size_type i = size_; i < new_size; ++i) {
                new(data_ + i) T();
            }
        } else if (new_size < size_) {
            // 析构多余元素
            for (size_type i = new_size; i < size_; ++i) {
                data_[i].~T();
            }
        }
        
        size_ = new_size;
    }
    
    void resize(size_type new_size, const T& value) {
        if (new_size > capacity_) {
            reserve(new_size);
        }
        
        if (new_size > size_) {
            // 用指定值构造新元素
            for (size_type i = size_; i < new_size; ++i) {
                new(data_ + i) T(value);
            }
        } else if (new_size < size_) {
            // 析构多余元素
            for (size_type i = new_size; i < size_; ++i) {
                data_[i].~T();
            }
        }
        
        size_ = new_size;
    }
    
    void shrink_to_fit() {
        if (size_ < capacity_) {
            reallocate(size_);
        }
    }
    
    // ==================== 元素访问 ====================
    
    reference operator[](size_type index) noexcept { return data_[index]; }
    const_reference operator[](size_type index) const noexcept { return data_[index]; }
    
    reference at(size_type index) {
        if (index >= size_) throw std::out_of_range("TXVector::at");
        return data_[index];
    }
    
    const_reference at(size_type index) const {
        if (index >= size_) throw std::out_of_range("TXVector::at");
        return data_[index];
    }
    
    reference front() noexcept { return data_[0]; }
    const_reference front() const noexcept { return data_[0]; }
    
    reference back() noexcept { return data_[size_ - 1]; }
    const_reference back() const noexcept { return data_[size_ - 1]; }
    
    T* data() noexcept { return data_; }
    const T* data() const noexcept { return data_; }
    
    // ==================== 迭代器 ====================
    
    iterator begin() noexcept { return data_; }
    const_iterator begin() const noexcept { return data_; }
    const_iterator cbegin() const noexcept { return data_; }
    
    iterator end() noexcept { return data_ + size_; }
    const_iterator end() const noexcept { return data_ + size_; }
    const_iterator cend() const noexcept { return data_ + size_; }
    
    // ==================== 修改操作 ====================
    
    void clear() noexcept {
        for (size_type i = 0; i < size_; ++i) {
            data_[i].~T();
        }
        size_ = 0;
    }
    
    void push_back(const T& value) {
        if (size_ >= capacity_) {
            reserve(calculateGrowth());
        }
        new(data_ + size_) T(value);
        ++size_;
    }
    
    void push_back(T&& value) {
        if (size_ >= capacity_) {
            reserve(calculateGrowth());
        }
        new(data_ + size_) T(std::move(value));
        ++size_;
    }
    
    template<typename... Args>
    reference emplace_back(Args&&... args) {
        if (size_ >= capacity_) {
            reserve(calculateGrowth());
        }
        new(data_ + size_) T(std::forward<Args>(args)...);
        return data_[size_++];
    }
    
    void pop_back() noexcept {
        if (size_ > 0) {
            data_[--size_].~T();
        }
    }
    
    // ==================== 批量操作 ====================
    
    void assign(size_type count, const T& value) {
        clear();
        reserve(count);
        for (size_type i = 0; i < count; ++i) {
            new(data_ + i) T(value);
        }
        size_ = count;
    }
    
    template<typename InputIt>
    void assign(InputIt first, InputIt last) {
        clear();
        size_type count = std::distance(first, last);
        reserve(count);
        
        size_type i = 0;
        for (auto it = first; it != last; ++it, ++i) {
            new(data_ + i) T(*it);
        }
        size_ = count;
    }
    
    void assign(std::initializer_list<T> init) {
        assign(init.begin(), init.end());
    }

private:
    // ==================== 内部辅助方法 ====================
    
    /**
     * @brief 🚀 SIMD友好的容量对齐
     */
    size_type alignCapacity(size_type capacity) const {
        if (capacity < MIN_CAPACITY) capacity = MIN_CAPACITY;
        
        // 对齐到SIMD边界
        size_type element_size = sizeof(T);
        size_type elements_per_simd = SIMD_ALIGNMENT / element_size;
        if (elements_per_simd > 0) {
            capacity = ((capacity + elements_per_simd - 1) / elements_per_simd) * elements_per_simd;
        }
        
        return capacity;
    }
    
    /**
     * @brief 计算增长容量
     */
    size_type calculateGrowth() const {
        return capacity_ == 0 ? MIN_CAPACITY : capacity_ * 2;
    }
    
    /**
     * @brief 🚀 高性能重新分配
     */
    void reallocate(size_type new_capacity) {
        if (new_capacity == 0) {
            deallocate();
            return;
        }
        
        // 🚀 使用内存管理器分配新内存
        size_type bytes_needed = new_capacity * sizeof(T);
        T* new_data = static_cast<T*>(memory_manager_->allocate(bytes_needed));

        if (!new_data) {
            throw std::bad_alloc();
        }
        
        // 移动现有元素
        if (data_ && size_ > 0) {
            for (size_type i = 0; i < size_; ++i) {
                new(new_data + i) T(std::move(data_[i]));
                data_[i].~T();
            }
        }
        
        // 释放旧内存
        if (data_) {
            memory_manager_->deallocate(data_);
        }
        
        data_ = new_data;
        capacity_ = new_capacity;
    }
    
    /**
     * @brief 释放内存
     */
    void deallocate() {
        if (data_) {
            memory_manager_->deallocate(data_);
            data_ = nullptr;
            capacity_ = 0;
        }
    }
};

} // namespace TinaXlsx
