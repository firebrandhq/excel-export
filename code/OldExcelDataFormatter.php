<?php

/**
 * OldExcelDataFormatter extends {@link ExcelDataFormatter} to provide a DataFormatter
 * suitable for exporting an {@link SS_link} of {@link DataObjectInterface} to
 * a Excel5 spreadsheet (XLS).
 *
 * @author Firebrand <hello@firebrand.nz>
 * @license MIT
 * @package silverstripe-excel-export
 */
class OldExcelDataFormatter extends ExcelDataFormatter
{

    /**
     * @inheritdoc
     */
    public function supportedExtensions()
    {
        return array(
            'xls',
        );
    }

    /**
     * @inheritdoc
     */
    public function supportedMimeTypes()
    {
        return array(
            'application/vnd.ms-excel',
        );
    }

    /**
     * @inheritdoc
     */
    public function convertDataObjectSet(SS_List $set)
    {
        $this->setHeader();

        $excel = $this->getPhpExcelObject($set);

        $fileData = $this->getFileData($excel, 'Excel5');

        return $fileData;
    }
}
