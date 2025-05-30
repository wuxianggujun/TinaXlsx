/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2025, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/xlsxwriter/styles.h"

// Test _xml_declaration().
CTEST(styles, xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_styles *styles = lxw_styles_new();
    styles->file = testfile;

    _styles_xml_declaration(styles);

    RUN_XLSX_STREQ(exp, got);

    lxw_styles_free(styles);
}
