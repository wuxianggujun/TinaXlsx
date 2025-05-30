//
// Created by wuxianggujun on 2025/5/30.
//
#pragma once

#include "TXError.hpp" // 您的 TXError.hpp 文件路径
#include <type_traits>
#include <stdexcept>
#include <functional>
#include <optional>
#include <string>
#include <type_traits>
#include <utility>
#include <variant>


namespace TinaXlsx
{
    // 前向声明 TXResult 类
    template <typename T>
    class [[nodiscard]] TXResult; // 为非void版本添加 [[nodiscard]]
    template <>
    class [[nodiscard]] TXResult<void>; // 为void特化版本添加 [[nodiscard]]


    /**
     * @brief 通用结果类型，用于表示成功值或错误。
     * @tparam T 成功时返回的类型。
     */
    template <typename T>
    class [[nodiscard]] TXResult
    {
    public:
        // 辅助类型，用于模板元编程
        using value_type = T;
        using error_type = TXError;

        // 成功时构造函数
        explicit TXResult(const T& value): data_(value)
        {
        }

        explicit TXResult(T&& value) noexcept(std::is_nothrow_move_constructible_v<T>): data_(std::move(value))
        {
        }

        // 错误时构造函数
        explicit TXResult(const TXError& error): data_(error)
        {
        }

        explicit TXResult(TXError&& error) noexcept(std::is_nothrow_move_constructible_v<TXError>): data_(
            std::move(error))
        {
        }

        // 从其他 TXResult 转换构造函数 (const& 版本)
        template <typename U, typename = std::enable_if_t<std::is_constructible_v<T, const U&>>>
        explicit TXResult(const TXResult<U>& other): data_{}
        {
            if (other.isOk()) // 使用 isOk()
            {
                data_.template emplace<T>(other.value());
            }
            else
            {
                data_.template emplace<TXError>(other.error());
            }
        }

        // 从其他 TXResult 转换构造函数 (&& 版本)
        template <typename U, typename = std::enable_if_t<std::is_constructible_v<T, U&&>>>
        explicit TXResult(TXResult<U>&& other) noexcept(
            std::is_nothrow_constructible_v<T, U&&> && std::is_nothrow_move_constructible_v<TXError>)
            : data_{}
        {
            if (other.isOk()) // 使用 isOk()
            {
                // std::move(other).value() 会调用 TXResult<U>::value() &&
                data_.template emplace<T>(std::move(other).value());
            }
            else
            {
                data_.template emplace<TXError>(std::move(other).error());
            }
        }

        // 默认拷贝和移动构造/赋值
        TXResult(const TXResult&) = default;
        TXResult(TXResult&&) noexcept = default;
        TXResult& operator=(const TXResult&) = default;
        TXResult& operator=(TXResult&&) noexcept = default;

        // ==================== 状态检查 ====================
        /**
         * @brief 检查是否包含有效值 (成功状态)。
         */
        [[nodiscard]] bool isOk() const noexcept
        {
            return std::holds_alternative<T>(data_);
        }

        /**
         * @brief 检查是否包含错误。
         */
        [[nodiscard]] bool isError() const noexcept
        {
            return std::holds_alternative<TXError>(data_);
        }

        /**
         * @brief 检查是否包含特定错误代码。
         */
        [[nodiscard]] bool hasErrorCode(TXErrorCode code) const noexcept
        {
            return isError() && std::get<TXError>(data_).getCode() == code;
        }

        // ==================== 值访问 ====================
        /**
         * @brief 获取值的引用 (左值引用限定)。如果当前是错误状态，则抛出异常。
         */
        [[nodiscard]] T& value() &
        {
            if (isOk())
            {
                return std::get<T>(data_);
            }
            throw std::runtime_error(
                "TXResult: Attempt to get value from error TXResult: " + std::get<TXError>(data_).toDetailString());
        }

        /**
         * @brief 获取值的常量引用 (const左值引用限定)。如果当前是错误状态，则抛出异常。
         */
        [[nodiscard]] const T& value() const &
        {
            if (isOk())
            {
                return std::get<T>(data_);
            }
            throw std::runtime_error(
                "TXResult: Attempt to get value from error TXResult: " + std::get<TXError>(data_).toDetailString());
        }

        /**
         * @brief 获取值的右值引用 (右值引用限定)。如果当前是错误状态，则抛出异常。
         */
        [[nodiscard]] T&& value() &&
        {
            if (isOk())
            {
                return std::move(std::get<T>(data_));
            }
            throw std::runtime_error( // 注意：这里移动错误信息，但错误信息本身通常是可拷贝的
                "TXResult: Attempt to get value from error TXResult: " + std::move(std::get<TXError>(data_)).toDetailString());
        }

        /**
         * @brief 获取值的常量右值引用 (const右值引用限定)。如果当前是错误状态，则抛出异常。
         */
        [[nodiscard]] const T&& value() const &&
        {
            if (isOk())
            {
                return std::move(std::get<T>(data_));
            }
            throw std::runtime_error(
                "TXResult: Attempt to get value from error TXResult: " + std::move(std::get<TXError>(data_)).toDetailString());
        }

        /**
         * @brief 获取值，如果当前是错误状态，则返回提供的默认值 (const左值引用限定)。
         */
        template <typename U>
        [[nodiscard]] T valueOr(U&& defaultValue) const &
        {
            return isOk() ? std::get<T>(data_) : static_cast<T>(std::forward<U>(defaultValue));
        }

        /**
         * @brief 获取值，如果当前是错误状态，则返回提供的默认值 (右值引用限定)。
         */
        template <typename U>
        [[nodiscard]] T valueOr(U&& defaultValue) &&
        {
            return isOk() ? std::move(std::get<T>(data_)) : static_cast<T>(std::forward<U>(defaultValue));
        }

        /**
         * @brief 获取错误的引用 (左值引用限定)。如果当前是成功状态，则抛出异常。
         */
        [[nodiscard]] TXError& error() &
        {
            if (isError()) return std::get<TXError>(data_);
            throw std::runtime_error("TXResult: Attempt to get error from successful TXResult");
        }
        /**
         * @brief 获取错误的常量引用 (const左值引用限定)。如果当前是成功状态，则抛出异常。
         */
        [[nodiscard]] const TXError& error() const &
        {
            if (isError()) return std::get<TXError>(data_);
            throw std::runtime_error("TXResult: Attempt to get error from successful TXResult");
        }
        /**
         * @brief 获取错误的右值引用 (右值引用限定)。如果当前是成功状态，则抛出异常。
         */
        [[nodiscard]] TXError&& error() &&
        {
            if (isError()) return std::move(std::get<TXError>(data_));
            throw std::runtime_error("TXResult: Attempt to get error from successful TXResult");
        }
        /**
         * @brief 获取错误的常量右值引用 (const右值引用限定)。如果当前是成功状态，则抛出异常。
         */
        [[nodiscard]] const TXError&& error() const &&
        {
            if (isError()) return std::move(std::get<TXError>(data_));
            throw std::runtime_error("TXResult: Attempt to get error from successful TXResult");
        }


        // ==================== 上下文与错误链 ====================
        /**
         * @brief 为错误结果追加上下文信息 (左值引用限定)。
         * @param context_message 要追加的上下文消息。
         * @return TXResult& 当前对象的引用，支持链式调用。
         * @note 依赖 TXError::appendContext 方法。
         */
        [[nodiscard]] TXResult& withContext(const std::string& context_message) &
        {
            if (isError())
            {
                std::get<TXError>(data_).appendContext(context_message);
            }
            return *this;
        }

        /**
         * @brief 为错误结果追加上下文信息 (右值引用限定)。
         */
        [[nodiscard]] TXResult&& withContext(const std::string& context_message) &&
        {
            if (isError())
            {
                // 对于右值操作，我们修改的是即将被移走的对象的内部状态
                std::get<TXError>(data_).appendContext(context_message);
            }
            return std::move(*this);
        }

        /**
         * @brief 包装当前错误（如果存在）为一个新的、更高级别的错误 (右值引用限定)。
         * 此方法会消耗（移动）当前的 TXResult 对象。
         * @param outer_code 新的外部错误的错误码。
         * @param outer_message 新的外部错误的描述信息。
         * @return 一个新的 TXResult 对象。
         */
        [[nodiscard]] TXResult<T> wrapError(TXErrorCode outer_code, const std::string& outer_message) &&
        {
            if (isError()) // 已修正：之前是 isSuccess()，逻辑错误
            {
                TXError outer_error(outer_code, outer_message);
                outer_error.setCause(std::move(std::get<TXError>(data_)));
                return TXResult<T>(std::move(outer_error));
            }
            return std::move(*this);
        }

        /**
         * @brief 包装当前错误（如果存在）为一个新的、更高级别的错误 (const左值引用限定)。
         * 此方法会拷贝当前的 TXResult 对象 (如果成功) 或其内部的 TXError 对象 (如果失败)。
         */
        [[nodiscard]] TXResult<T> wrapError(TXErrorCode outer_code, const std::string& outer_message) const &
        {
            if (isError()) // 已修正：之前是 isSuccess()，逻辑错误
            {
                TXError outer_error(outer_code, outer_message);
                outer_error.setCause(std::get<TXError>(data_)); // 拷贝内部错误作为 cause
                return TXResult<T>(std::move(outer_error));
            }
            return *this; // 如果是 Ok，则拷贝返回
        }

        // ==================== 链式操作 (Monadic Operations) ====================
        
        // --- map ---
        template <typename Fn> // Fn: T& -> U  (或 Fn: T& -> void)
        [[nodiscard]] auto map(Fn&& fn) & -> TXResult<std::invoke_result_t<Fn&, T&>>
        {
            using ReturnType = std::invoke_result_t<Fn&, T&>;
            if (isOk())
            {
                if constexpr (std::is_void_v<ReturnType>)
                {
                    std::invoke(std::forward<Fn>(fn), value());
                    return TXResult<void>(); 
                }
                else
                {
                    return TXResult<ReturnType>(std::invoke(std::forward<Fn>(fn), value()));
                }
            }
            return TXResult<ReturnType>(error()); // 传播错误 (拷贝错误对象)
        }

        template <typename Fn> // Fn: const T& -> U (或 Fn: const T& -> void)
        [[nodiscard]] auto map(Fn&& fn) const & -> TXResult<std::invoke_result_t<Fn&, const T&>>
        {
            using ReturnType = std::invoke_result_t<Fn&, const T&>;
            if (isOk())
            {
                if constexpr (std::is_void_v<ReturnType>)
                {
                    std::invoke(std::forward<Fn>(fn), value());
                    return TXResult<void>();
                }
                else
                {
                    return TXResult<ReturnType>(std::invoke(std::forward<Fn>(fn), value()));
                }
            }
            return TXResult<ReturnType>(error());
        }

        template <typename Fn> // Fn: T&& -> U (或 Fn: T&& -> void)
        [[nodiscard]] auto map(Fn&& fn) && -> TXResult<std::invoke_result_t<Fn&&, T&&>>
        {
            using ReturnType = std::invoke_result_t<Fn&&, T&&>;
            if (isOk())
            {
                if constexpr (std::is_void_v<ReturnType>)
                {
                    std::invoke(std::forward<Fn>(fn), std::move(*this).value());
                    return TXResult<void>();
                }
                else
                {
                    return TXResult<ReturnType>(std::invoke(std::forward<Fn>(fn), std::move(*this).value()));
                }
            }
            return TXResult<ReturnType>(std::move(*this).error()); // 传播错误 (移动错误对象)
        }
        
        template <typename Fn> // Fn: const T&& -> U (或 Fn: const T&& -> void)
        [[nodiscard]] auto map(Fn&& fn) const && -> TXResult<std::invoke_result_t<Fn&&, const T&&>>
        {
            using ReturnType = std::invoke_result_t<Fn&&, const T&&>;
            if (isOk())
            {
                if constexpr (std::is_void_v<ReturnType>)
                {
                    std::invoke(std::forward<Fn>(fn), std::move(*this).value());
                    return TXResult<void>();
                }
                else
                {
                    return TXResult<ReturnType>(std::invoke(std::forward<Fn>(fn), std::move(*this).value()));
                }
            }
            return TXResult<ReturnType>(std::move(*this).error());
        }


        // --- andThen ---
        /**
         * @brief 如果成功，则对成功值应用函数 `fn`，`fn` 必须返回一个 `TXResult`；否则，传播错误。
         */
        template <typename Fn> // Fn: T& -> TXResult<U> (或 Fn: T& -> TXResult<void>)
        [[nodiscard]] auto andThen(Fn&& fn) & -> std::invoke_result_t<Fn&, T&>
        {
            using NextResultType = std::invoke_result_t<Fn&, T&>;
            if (isOk())
            {
                try { return std::invoke(std::forward<Fn>(fn), value()); }
                catch (const std::exception& e) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Exception in andThen callback: " + std::string(e.what()))); }
                catch (...) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Unknown exception in andThen callback"));}
            }
            return NextResultType(error()); // 传播原错误 (拷贝错误对象)
        }
        
        template <typename Fn> // Fn: const T& -> TXResult<U>
        [[nodiscard]] auto andThen(Fn&& fn) const & -> std::invoke_result_t<Fn&, const T&>
        {
            using NextResultType = std::invoke_result_t<Fn&, const T&>;
            if (isOk())
            {
                try { return std::invoke(std::forward<Fn>(fn), value()); }
                catch (const std::exception& e) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Exception in andThen callback: " + std::string(e.what()))); }
                catch (...) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Unknown exception in andThen callback"));}
            }
            return NextResultType(error());
        }

        template <typename Fn> // Fn: T&& -> TXResult<U>
        [[nodiscard]] auto andThen(Fn&& fn) && -> std::invoke_result_t<Fn&&, T&&>
        {
            using NextResultType = std::invoke_result_t<Fn&&, T&&>;
            if (isOk())
            {
                try { return std::invoke(std::forward<Fn>(fn), std::move(*this).value()); }
                catch (const std::exception& e) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Exception in andThen callback: " + std::string(e.what()))); }
                catch (...) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Unknown exception in andThen callback"));}
            }
            return NextResultType(std::move(*this).error()); // 传播原错误 (移动错误对象)
        }
        
        template <typename Fn> // Fn: const T&& -> TXResult<U>
        [[nodiscard]] auto andThen(Fn&& fn) const && -> std::invoke_result_t<Fn&&, const T&&>
        {
            using NextResultType = std::invoke_result_t<Fn&&, const T&&>;
            if (isOk())
            {
                try { return std::invoke(std::forward<Fn>(fn), std::move(*this).value()); }
                catch (const std::exception& e) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Exception in andThen callback: " + std::string(e.what()))); }
                catch (...) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Unknown exception in andThen callback"));}
            }
            return NextResultType(std::move(*this).error());
        }


        // --- orElse ---
        /**
         * @brief 如果错误，则对错误值应用函数 `fn`；否则，传播成功结果。
         * `fn` 通常用于恢复错误，并且必须返回 TXResult<T> (与原类型相同的 T)。
         */
        template <typename Fn> // Fn: TXError& -> TXResult<T>
        [[nodiscard]] TXResult<T> orElse(Fn&& fn) & 
        {
            static_assert(std::is_same_v<std::invoke_result_t<Fn&, TXError&>, TXResult<T>>,
                          "orElse's recovery function must return a TXResult<T> of the original T type.");
            if (isError())
            {
                return std::invoke(std::forward<Fn>(fn), error());
            }
            return *this; // 返回当前对象的拷贝 (因为是左值引用限定)
        }
        
        template <typename Fn> // Fn: const TXError& -> TXResult<T>
        [[nodiscard]] TXResult<T> orElse(Fn&& fn) const &
        {
            static_assert(std::is_same_v<std::invoke_result_t<Fn&, const TXError&>, TXResult<T>>,
                          "orElse's recovery function must return a TXResult<T> of the original T type.");
            if (isError())
            {
                return std::invoke(std::forward<Fn>(fn), error());
            }
            return *this; 
        }
        
        template <typename Fn> // Fn: TXError&& -> TXResult<T>
        [[nodiscard]] TXResult<T> orElse(Fn&& fn) &&
        {
            static_assert(std::is_same_v<std::invoke_result_t<Fn&&, TXError&&>, TXResult<T>>,
                          "orElse's recovery function must return a TXResult<T> of the original T type.");
            if (isError())
            {
                return std::invoke(std::forward<Fn>(fn), std::move(*this).error());
            }
            return std::move(*this); // 移动当前对象
        }
        
        template <typename Fn> // Fn: const TXError&& -> TXResult<T>
        [[nodiscard]] TXResult<T> orElse(Fn&& fn) const &&
        {
            static_assert(std::is_same_v<std::invoke_result_t<Fn&&, const TXError&&>, TXResult<T>>,
                          "orElse's recovery function must return a TXResult<T> of the original T type.");
            if (isError())
            {
                return std::invoke(std::forward<Fn>(fn), std::move(*this).error());
            }
            return std::move(*this);
        }


        // ==================== 运算符重载 ====================
        [[nodiscard]] explicit operator bool() const noexcept
        {
            return isOk();
        }

        [[nodiscard]] T* operator->() { return &value(); }
        [[nodiscard]] const T* operator->() const { return &value(); }
        [[nodiscard]] T& operator*() & { return value(); }
        [[nodiscard]] const T& operator*() const & { return value(); }
        [[nodiscard]] T&& operator*() && { return std::move(value()); }
        [[nodiscard]] const T&& operator*() const && { return std::move(value()); }

    private:
        std::variant<T, TXError> data_;
    };


    /**
     * @brief void类型的Result特化
     */
    template <>
    class [[nodiscard]] TXResult<void>
    {
    public:
        using value_type = void; // 辅助类型
        using error_type = TXError;

        // 成功构造函数
        TXResult() noexcept : error_(std::nullopt) {}

        // 错误构造函数
        explicit TXResult(const TXError& error) : error_(error) {}
        explicit TXResult(TXError&& error) noexcept(std::is_nothrow_move_constructible_v<TXError>) : error_(std::move(error)){}

        // 默认拷贝和移动
        TXResult(const TXResult&) = default;
        TXResult(TXResult&&) noexcept = default;
        TXResult& operator=(const TXResult&) = default;
        TXResult& operator=(TXResult&&) noexcept = default;

        // ==================== 状态检查 ====================
        [[nodiscard]] bool isOk() const noexcept { return !error_.has_value(); }
        [[nodiscard]] bool isError() const noexcept { return error_.has_value(); }

        [[nodiscard]] bool hasErrorCode(TXErrorCode code) const noexcept
        {
            return isError() && error_.value().getCode() == code;
        }

        // ==================== 值访问 ====================
        void value() const 
        {
            if (isError())
            {
                throw std::runtime_error(
                    "TXResult<void>: Attempt to get value from error TXResult: " + error_.value().toDetailString());
            }
            // 成功时无操作
        }

        [[nodiscard]] TXError& error() &
        {
            if (isError()) return *error_; // error_ 是 std::optional, *error_ 获取其值
            throw std::runtime_error("TXResult<void>: Attempt to get error from successful TXResult");
        }
        [[nodiscard]] const TXError& error() const &
        {
            if (isError()) return *error_;
            throw std::runtime_error("TXResult<void>: Attempt to get error from successful TXResult");
        }
        [[nodiscard]] TXError&& error() &&
        {
            if (isError()) return std::move(*error_);
            throw std::runtime_error("TXResult<void>: Attempt to get error from successful TXResult");
        }
        [[nodiscard]] const TXError&& error() const &&
        {
            if (isError()) return std::move(*error_);
            throw std::runtime_error("TXResult<void>: Attempt to get error from successful TXResult");
        }

        // ==================== 上下文与错误链 ====================
        [[nodiscard]] TXResult<void>& withContext(const std::string& context_message) &
        {
            if (isError())
            {
                error_->appendContext(context_message); 
            }
            return *this;
        }
        [[nodiscard]] TXResult<void>&& withContext(const std::string& context_message) &&
        {
            if (isError())
            {
                error_->appendContext(context_message);
            }
            return std::move(*this);
        }

        [[nodiscard]] TXResult<void> wrapError(TXErrorCode outer_code, const std::string& outer_message) &&
        {
            if (isError())
            {
                TXError outer_err(outer_code, outer_message);
                outer_err.setCause(std::move(*error_)); 
                return TXResult<void>(std::move(outer_err));
            }
            return std::move(*this);
        }
        [[nodiscard]] TXResult<void> wrapError(TXErrorCode outer_code, const std::string& outer_message) const &
        {
            if (isError())
            {
                TXError outer_err(outer_code, outer_message);
                outer_err.setCause(*error_); // 拷贝
                return TXResult<void>(std::move(outer_err));
            }
            return *this;
        }

   // const & 版本
template <typename Fn>
[[nodiscard]] auto map(Fn&& fn) const & -> TXResult<std::invoke_result_t<Fn&>>
{
    using ReturnType = std::invoke_result_t<Fn&>;
    if (isOk())
    {
        if constexpr (std::is_void_v<ReturnType>)
        {
            std::invoke(std::forward<Fn>(fn));
            return TXResult<void>();
        }
        else
        {
            return TXResult<ReturnType>(std::invoke(std::forward<Fn>(fn)));
        }
    }
    return TXResult<ReturnType>(error()); // error() 是 const &
}

// && 版本 (允许移动 error_)
template <typename Fn>
[[nodiscard]] auto map(Fn&& fn) && -> TXResult<std::invoke_result_t<Fn&&>>
{
    using ReturnType = std::invoke_result_t<Fn&&>;
    if (isOk())
    {
        if constexpr (std::is_void_v<ReturnType>)
        {
            std::invoke(std::forward<Fn>(fn));
            return TXResult<void>();
        }
        else
        {
            return TXResult<ReturnType>(std::invoke(std::forward<Fn>(fn)));
        }
    }
    return TXResult<ReturnType>(std::move(*this).error()); // error() 是 &&
}

     // const & 版本
template <typename Fn>
[[nodiscard]] auto andThen(Fn&& fn) const & -> std::invoke_result_t<Fn&>
{
    using NextResultType = std::invoke_result_t<Fn&>;
    if (isOk())
    {
        try { return std::invoke(std::forward<Fn>(fn)); }
        catch (const std::exception& e) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Exception in andThen<void> const& callback: " + std::string(e.what()))); }
        catch (...) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Unknown exception in andThen<void> const& callback"));}
    }
    return NextResultType(error()); // error() 是 const &
}

// && 版本
template <typename Fn>
[[nodiscard]] auto andThen(Fn&& fn) && -> std::invoke_result_t<Fn&&>
{
    using NextResultType = std::invoke_result_t<Fn&&>;
    if (isOk())
    {
        try { return std::invoke(std::forward<Fn>(fn)); }
        catch (const std::exception& e) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Exception in andThen<void> && callback: " + std::string(e.what()))); }
        catch (...) { return NextResultType(TXError(TXErrorCode::OperationFailed, "Unknown exception in andThen<void> && callback"));}
    }
    return NextResultType(std::move(*this).error()); // error() 是 &&
}

        // --- orElse for void ---
        template <typename Fn> // Fn: TXError& -> TXResult<void>
        [[nodiscard]] TXResult<void> orElse(Fn&& fn) &
        {
            static_assert(std::is_same_v<std::invoke_result_t<Fn&, TXError&>, TXResult<void>>,
                          "orElse's recovery function for TXResult<void> must return TXResult<void>.");
            if (isError())
            {
                return std::invoke(std::forward<Fn>(fn), error());
            }
            return *this; 
        }
        
        template <typename Fn> // Fn: const TXError& -> TXResult<void>
        [[nodiscard]] TXResult<void> orElse(Fn&& fn) const &
        {
            static_assert(std::is_same_v<std::invoke_result_t<Fn&, const TXError&>, TXResult<void>>,
                          "orElse's recovery function for TXResult<void> must return TXResult<void>.");
            if (isError())
            {
                return std::invoke(std::forward<Fn>(fn), error());
            }
            return *this; 
        }
        
        template <typename Fn> // Fn: TXError&& -> TXResult<void>
        [[nodiscard]] TXResult<void> orElse(Fn&& fn) &&
        {
            static_assert(std::is_same_v<std::invoke_result_t<Fn&&, TXError&&>, TXResult<void>>,
                          "orElse's recovery function for TXResult<void> must return TXResult<void>.");
            if (isError())
            {
                return std::invoke(std::forward<Fn>(fn), std::move(*this).error());
            }
            return std::move(*this); 
        }
        
        template <typename Fn> // Fn: const TXError&& -> TXResult<void>
        [[nodiscard]] TXResult<void> orElse(Fn&& fn) const &&
        {
            static_assert(std::is_same_v<std::invoke_result_t<Fn&&, const TXError&&>, TXResult<void>>,
                          "orElse's recovery function for TXResult<void> must return TXResult<void>.");
            if (isError())
            {
                return std::invoke(std::forward<Fn>(fn), std::move(*this).error());
            }
            return std::move(*this);
        }


        // ==================== 运算符重载 ====================
        [[nodiscard]] explicit operator bool() const noexcept { return isOk(); }

    private:
        // `std::nullopt` 表示成功 (Ok)
        std::optional<TXError> error_; 
    };


    // ==================== 辅助工厂函数 ====================
    /**
     * @brief 创建成功结果 (非void类型)。
     */
    template <typename T>
    [[nodiscard]] inline TXResult<T> Ok(T&& value)
    {
        return TXResult<T>(std::forward<T>(value));
    }
    /**
     * @brief 创建成功结果 (非void类型, const 左值引用版本)。
     */
    template <typename T>
    [[nodiscard]] inline TXResult<T> Ok(const T& value)
    {
        return TXResult<T>(value);
    }

    /**
     * @brief 创建成功结果 (void类型)。
     */
    [[nodiscard]] inline TXResult<void> Ok()
    {
        return TXResult<void>();
    }

    /**
     * @brief 创建错误结果 (目标T类型)。
     * @param error TXError 对象。
     */
    template <typename T> 
    [[nodiscard]] inline TXResult<T> Err(TXError error) // 接受 TXError 对象 (会被移动或拷贝)
    {
        return TXResult<T>(std::move(error));
    }

    /**
     * @brief 创建错误结果 (目标T类型)，通过错误码和消息。
     */
    template <typename T> 
    [[nodiscard]] inline TXResult<T> Err(TXErrorCode code, const std::string& message = "")
    {
        return TXResult<T>(TXError(code, message));
    }

    /**
     * @brief 创建错误结果 (void类型)。
     */
    [[nodiscard]] inline TXResult<void> Err(TXError error) // 接受 TXError 对象
    {
        return TXResult<void>(std::move(error));
    }
    /**
     * @brief 创建错误结果 (void类型)，通过错误码和消息。
     */
    [[nodiscard]] inline TXResult<void> Err(TXErrorCode code, const std::string& message = "")
    {
        return TXResult<void>(TXError(code, message));
    }

} // namespace TinaXlsx
