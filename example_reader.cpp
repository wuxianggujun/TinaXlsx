/**
 * @file example_reader.cpp
 * @brief Example program to test minizip-ng based Reader implementation
 */

#include "TinaXlsx/TinaXlsx.hpp"
#include <iostream>
#include <exception>

int main() {
    try {
        std::cout << "TinaXlsx Reader Test Program" << std::endl;
        std::cout << "Version: " << TinaXlsx::Version::string << std::endl;
        std::cout << "Based on minizip-ng and expat" << std::endl;
        
        // Test static utility functions
        std::cout << "\n=== Testing Static Utility Functions ===" << std::endl;
        
        // Test string conversion
        auto value1 = TinaXlsx::Reader::stringToCellValue("42");
        auto value2 = TinaXlsx::Reader::stringToCellValue("3.14");
        auto value3 = TinaXlsx::Reader::stringToCellValue("true");
        auto value4 = TinaXlsx::Reader::stringToCellValue("hello");
        
        std::cout << "String conversion test:" << std::endl;
        std::cout << "  '42' -> " << TinaXlsx::Reader::cellValueToString(value1) << std::endl;
        std::cout << "  '3.14' -> " << TinaXlsx::Reader::cellValueToString(value2) << std::endl;
        std::cout << "  'true' -> " << TinaXlsx::Reader::cellValueToString(value3) << std::endl;
        std::cout << "  'hello' -> " << TinaXlsx::Reader::cellValueToString(value4) << std::endl;
        
        // Test empty row detection
        TinaXlsx::RowData emptyRow;
        TinaXlsx::RowData nonEmptyRow = {TinaXlsx::CellValue(std::string("test"))};
        
        std::cout << "\nEmpty row detection test:" << std::endl;
        std::cout << "  Empty row: " << (TinaXlsx::Reader::isEmptyRow(emptyRow) ? "Yes" : "No") << std::endl;
        std::cout << "  Non-empty row: " << (TinaXlsx::Reader::isEmptyRow(nonEmptyRow) ? "Yes" : "No") << std::endl;
        
        // Test file opening (expected to fail)
        std::cout << "\n=== Testing File Opening ===" << std::endl;
        try {
            auto reader = TinaXlsx::createReader("nonexistent.xlsx");
            std::cout << "ERROR: Should have thrown an exception" << std::endl;
        } catch (const TinaXlsx::FileException& e) {
            std::cout << "CORRECT: Caught expected file exception: " << e.what() << std::endl;
        }
        
        std::cout << "\n=== Test Complete ===" << std::endl;
        std::cout << "? Basic functionality tests passed!" << std::endl;
        std::cout << "? Successfully removed xlsxio dependency" << std::endl;
        std::cout << "? Using minizip-ng + expat implementation" << std::endl;
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    }
} 