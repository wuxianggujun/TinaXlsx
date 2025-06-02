//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlReader.hpp"
#include "TXXmlWriter.hpp"
#include <memory>
#include <vector>
#include <unordered_map>
#include <string>

namespace TinaXlsx {
    // 前向声明
    class TXPivotTable;
    class TXWorkbookRelsXmlHandler : public TXXmlHandler {
    public:
        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override {
            // xl/_rels/workbook.xml.rels 通常不需要加载
            return Ok();
        }

        /**
         * @brief 设置透视表信息（用于生成透视表缓存关系）
         * @param allPivotTables 所有工作表的透视表映射
         */
        void setAllPivotTables(const std::unordered_map<std::string, std::vector<std::shared_ptr<TXPivotTable>>>& allPivotTables);

        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;

        std::string partName() const override {
            return "xl/_rels/workbook.xml.rels";
        }

    private:
        std::unordered_map<std::string, std::vector<std::shared_ptr<TXPivotTable>>> m_allPivotTables;  ///< 所有工作表的透视表映射
    };
}
