//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include <string>
#include <memory>
#include "TinaXlsx/TXZipArchive.hpp"
#include "TinaXlsx/TXWorkbookContext.hpp"

namespace TinaXlsx
{
    class TXXmlHandler
    {
    public:
        virtual ~TXXmlHandler() = default;

        virtual bool load(TXZipArchiveReader& zipReader,TXWorkbookContext& context) = 0;

        virtual bool save(TXZipArchiveWriter& zipWriter,const TXWorkbookContext& context)  = 0;

        [[nodiscard]] virtual std::string_view partName() const = 0;

        [[nodiscard]] std::string lastError() const { return m_lastError; }

    protected:
        std::string m_lastError;
    };
}
