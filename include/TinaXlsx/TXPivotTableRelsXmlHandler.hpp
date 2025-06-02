//
// Created by wuxianggujun on 2025/1/15.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlWriter.hpp"

namespace TinaXlsx
{
    /**
     * @brief 透视表关系XML处理器
     * 
     * 处理透视表的关系文件，连接透视表定义和缓存定义
     */
    class TXPivotTableRelsXmlHandler : public TXXmlHandler
    {
    public:
        explicit TXPivotTableRelsXmlHandler(int pivotTableId);

        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override;
        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;
        [[nodiscard]] std::string partName() const override;

    private:
        int m_pivotTableId;  ///< 透视表ID
    };

} // namespace TinaXlsx
