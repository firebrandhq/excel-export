<?php

class CsvDataFormatter extends ExcelDataFormatter
{

    public function supportedExtensions()
    {
        return array(
            'csv',
        );
    }

    public function supportedMimeTypes()
    {
        return array(
            'text/csv',
        );
    }


    public function convertDataObjectSet(SS_List $set)
    {
        $this->setHeader();

        $excel = $this->getPhpExcelObject($set);

        $fileData = $this->getFileData($excel, 'CSV');

        return $fileData;
    }

}
