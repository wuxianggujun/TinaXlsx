//
// Created by wuxianggujun on 2025/1/15.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlWriter.hpp"

namespace TinaXlsx
{
    /**
     * @brief 工作表关系XML处理器
     * 
     * 处理工作表的关系文件，包括绘图关系
     */
    class TXWorksheetRelsXmlHandler : public TXXmlHandler
    {
    public:
        explicit TXWorksheetRelsXmlHandler(u32 sheetIndex);

        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override;
        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;
        [[nodiscard]] std::string partName() const override;

    private:
        u32 m_sheetIndex;
    };

    /**
     * @brief 绘图关系XML处理器
     * 
     * 处理绘图的关系文件，包括图表关系
     */
    class TXDrawingRelsXmlHandler : public TXXmlHandler
    {
    public:
        explicit TXDrawingRelsXmlHandler(u32 sheetIndex);

        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override;
        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;
        [[nodiscard]] std::string partName() const override;

    private:
        u32 m_sheetIndex;
    };

} // namespace TinaXlsx
