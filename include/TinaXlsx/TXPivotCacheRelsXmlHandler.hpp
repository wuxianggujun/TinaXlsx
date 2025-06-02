//
// Created by wuxianggujun on 2025/1/15.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlWriter.hpp"

namespace TinaXlsx
{
    /**
     * @brief 透视表缓存关系XML处理器
     * 
     * 处理透视表缓存的关系文件，连接缓存定义和缓存记录
     */
    class TXPivotCacheRelsXmlHandler : public TXXmlHandler
    {
    public:
        explicit TXPivotCacheRelsXmlHandler(int cacheId);

        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override;
        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;
        [[nodiscard]] std::string partName() const override;

    private:
        int m_cacheId;  ///< 缓存ID
    };

} // namespace TinaXlsx
