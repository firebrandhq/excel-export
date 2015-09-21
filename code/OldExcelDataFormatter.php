<?php

class OldExcelDataFormatter extends ExcelDataFormatter
{

    public function supportedExtensions()
    {
        return array(
            'xls',
        );
    }

    public function supportedMimeTypes()
    {
        return array(
            'application/vnd.ms-excel',
        );
    }


    public function convertDataObjectSet(SS_List $set)
    {
        $this->setHeader();

        $excel = $this->getPhpExcelObject($set);

        $fileData = $this->getFileData($excel, 'Excel5');

        return $fileData;
    }

}
