<?php

class ExcelDataFormatter extends DataFormatter
{

    private static $api_base = "api/v1/";

    public function supportedExtensions()
    {
        return array(
            'xlsx',
        );
    }

    public function supportedMimeTypes()
    {
        return array(
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        );
    }

    public function convertDataObject(DataObjectInterface $do)
    {
        return $this->convertDataObjectSet(new ArrayList(array($do)));
    }

    public function convertDataObjectSet(SS_List $set)
    {
        $this->setHeader();

        $excel = $this->getPhpExcelObject($set);

        $fileData = $this->getFileData($excel, 'Excel2007');

        return $fileData;
    }

    protected function setHeader()
    {
        Controller::curr()->getResponse()
            ->addHeader("Content-Type", $this->supportedMimeTypes()[0]);
    }

    protected function getFieldsForObj($do)
    {
        $fields = parent::getFieldsForObj($do);

        // Make sure our ID field is the first one.
        $fields = array('ID' => $fields['ID']) + $fields;

        return $fields;
    }


    public function getPhpExcelObject(SS_List $set) {
        $first = $set->first();

        // Get the Excel object
        $excel = $this->setupExcel($first);
        $sheet = $excel->setActiveSheetIndex(0);

        if ($first) {
            // Set up the header row
            $fields = $this->getFieldsForObj($first);
            $this->headerRow($sheet, $fields);

            foreach ($set as $item) {
                $this->addRow($sheet, $item, $fields);
            }

            // Freezing the first column and the header row
            $sheet->freezePane("B2");

            // Auto sizing all the columns
            $col = sizeof($fields);
            for ($i = 0; $i < $col; $i++) {
                $sheet
                    ->getColumnDimension(
                        PHPExcel_Cell::stringFromColumnIndex($i)
                    )
                    ->setAutoSize(true);
            }

        }

        return $excel;
    }

    protected function setupExcel(DataObjectInterface $do)
    {
        // Try to get the current user
        $member = Member::currentUser();
        $creator = $member ? $member->getName() : '';

        // Get information about the current Model Class
        $singular = $do ? $do->i18n_singular_name() : '';
        $plural = $do ? $do->i18n_plural_name() : '';

        // Create the Spread sheet
        $excel = new PHPExcel();

        $excel->getProperties()
            ->setCreator($creator)
            ->setTitle(_t(
                'firebrandhq.EXCELEXPORT',
                '{singular} export',
                'Title for the spread sheet export',
                array('singular' => $singular)
            ))
            ->setDescription(_t(
                'firebrandhq.EXCELEXPORT',
                'List of {plural} exported out of a SilverStripe website',
                'Description for the spread sheet export',
                array('pluralr' => $plural)
            ));

        $excel->getActiveSheet()->setTitle(_t(
            'firebrandhq.EXCELEXPORT',
            'Export'
        ));

        return $excel;
    }

    protected function headerRow(PHPExcel_Worksheet &$sheet, array $fields)
    {
        $row = 1;
        $col = 0;

        foreach ($fields as $field => $type) {
            $sheet->setCellValueByColumnAndRow($col, $row, $field);
            $col++;
        }

        // Get the last column
        $col--;
        $endcol = PHPExcel_Cell::stringFromColumnIndex($col);

        // Set Autofilters and Header row style
        $sheet->setAutoFilter("A1:{$endcol}1");
        $sheet->getStyle("A1:{$endcol}1")->getFont()->setBold(true);


        return $sheet;
    }

    protected function addRow(PHPExcel_Worksheet &$sheet, DataObjectInterface $item, array $fields)
    {
        $row = $sheet->getHighestRow() + 1;
        $col = 0;

        foreach ($fields as $field => $type) {
            $sheet->setCellValueByColumnAndRow($col, $row, $item->$field);
            $col++;
        }


        return $sheet;
    }

    protected function getFileData(PHPExcel $excel, $format) {
        $writer = PHPExcel_IOFactory::createWriter($excel, $format);
        ob_start();
        $writer->save('php://output');
        $fileData = ob_get_clean();

        return $fileData;
    }
}
