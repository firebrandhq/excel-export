<?php

/**
 * CsvDataFormatter extends {@link ExcelDataFormatter} to provide a DataFormatter
 * suitable for exporting an {@link SS_link} of {@link DataObjectInterface} to
 * a CSV spreadsheet.
 *
 * @author Firebrand <hello@firebrand.nz>
 * @license MIT
 * @package silverstripe-excel-export
 */
class CsvDataFormatter extends ExcelDataFormatter
{

    /**
     * @inheritdoc
     */
    public function supportedExtensions()
    {
        return array(
            'csv',
        );
    }

    /**
     * @inheritdoc
     */
    public function supportedMimeTypes()
    {
        return array(
            'text/csv',
        );
    }

    /**
     * @inheritdoc
     */
    public function convertDataObjectSet(SS_List $set)
    {
        $this->setHeader();

        $excel = $this->getPhpExcelObject($set);

        $fileData = $this->getFileData($excel, 'CSV');

        return $fileData;
    }
}
