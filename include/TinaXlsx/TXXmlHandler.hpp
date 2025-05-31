//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include <string>
#include <memory>
#include "TXZipArchive.hpp"
#include "TXWorkbookContext.hpp"
#include "TXResult.hpp"

namespace TinaXlsx
{
    class TXXmlHandler
    {
    public:
        virtual ~TXXmlHandler() = default;

        virtual TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) = 0;

        virtual TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) = 0;

        [[nodiscard]] virtual std::string partName() const = 0;

        [[nodiscard]] std::string lastError() const { return m_lastError; }

    protected:
        std::string m_lastError;
    };
}
