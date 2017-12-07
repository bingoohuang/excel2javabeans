package com.github.bingoohuang.excel2maps.impl;

import com.github.bingoohuang.excel2maps.ColumnDef;
import lombok.AllArgsConstructor;
import lombok.Value;
import lombok.val;
import org.apache.commons.lang3.StringUtils;

import java.util.Map;

@AllArgsConstructor @Value public class ColumnRef {
    ColumnDef columnDef;
    int columnIndex;

    public Ignored putMapOrIgnored(Map<String, String> map, String cellValue) {
        val upper = StringUtils.upperCase(cellValue);
        if (wildMatch(upper, columnDef.getIgnorePattern())) {
            return Ignored.YES;
        }

        map.put(columnDef.getColumnName(), cellValue);
        return Ignored.NO;
    }

    /**
     * The following Java method tests if a string matches a wildcard expression
     * (supporting ? for exactly one character or * for an arbitrary number of characters):
     *
     * @param text    Text to test
     * @param pattern (Wildcard) pattern to test
     * @return True if the text matches the wildcard pattern
     */
    public static boolean wildMatch(String text, String pattern) {
        if (pattern == null) return false;

        return text.matches(
                pattern.replace("?", ".?")
                        .replace("*", ".*?"));
    }
}
