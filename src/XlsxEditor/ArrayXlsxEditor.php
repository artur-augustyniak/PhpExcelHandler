<?php

namespace XlsxEditor;

/**
 * Description of ArrayXlsxEditor
 *
 * @author aaugustyniak
 */
class ArrayXlsxEditor extends XlsxEditor {

    public function writeFromDataSource() {
        $i = 1;
        foreach ($this->dataSource as $k => $v) {
            $keyCellCoord = 'A' . $i;
            $varCellCoord = 'B' . $i;
            $this->activeSheet->setCellValue($keyCellCoord, $k);
            $this->activeSheet->setCellValue($varCellCoord, $v);
            $i++;
        }
    }

}
