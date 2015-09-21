<?php
class GridFieldExcelExportButton implements
    GridField_HTMLProvider,
    GridField_ActionProvider,
    GridField_URLHandler
{

    /**
     * Fragment to write the button to
     */
    protected $targetFragment;

    public function __construct($targetFragment = "before")
    {
        $this->targetFragment = $targetFragment;
    }

    /**
     * Place the export button in a <p> tag below the field
     */
    public function getHTMLFragments($gridField)
    {
        $splitButton = new SplitButton('Export', 'Export');
        $splitButton->setAttribute('data-icon', 'download-csv');


        $button = new GridField_FormAction(
            $gridField,
            'xlsxexport',
            _t('firebrandhq.EXCELEXPORT', 'Export to Excel (XLSX)'),
            'xlsxexport',
            null
        );
        $button->addExtraClass('no-ajax');
        $splitButton->push($button);

        $button = new GridField_FormAction(
            $gridField,
            'xlsexport',
            _t('firebrandhq.EXCELEXPORT', 'Export to Excel (XLS)'),
            'xlsexport',
            null
        );
        $button->addExtraClass('no-ajax');
        $splitButton->push($button);

        $button = new GridField_FormAction(
            $gridField,
            'csvexport',
            _t('firebrandhq.EXCELEXPORT', 'Export to CSV'),
            'csvexport',
            null
        );
        $button->addExtraClass('no-ajax');

        $splitButton->push($button);

        return array(
            $this->targetFragment =>

                 $splitButton->Field()

        );
    }

    /**
     * export is an action button
     */
    public function getActions($gridField)
    {
        return array('xlsxexport', 'xlsexport', 'csvexport');
    }

    public function handleAction(
        GridField $gridField,
        $actionName,
        $arguments,
        $data
    ) {
        if ($actionName == 'xlsxexport') {
            return $this->handleXlsx($gridField);
        }

        if ($actionName == 'xlsexport') {
            return $this->handleXls($gridField);
        }

        if ($actionName == 'csvexport') {
            return $this->handleCsv($gridField);
        }
    }

    /**
     * it is also a URL
     */
    public function getURLHandlers($gridField)
    {
        return array(
            'xlsxexport' => 'handleXlsx',
            'xlsexport' => 'handleXls',
            'csvexport' => 'handleCsv',
        );
    }

    /**
     * Handle the export, for both the action button and the URL
      */
    public function handleXlsx(GridField $gridField, $request = null)
    {
        $items = $this->getItems($gridField);

        $this->setHeader($gridField, 'xlsx');

        $formater = new ExcelDataFormatter();
        $fileData = $formater->convertDataObjectSet($items);

        return $fileData;

        // return SS_HTTPRequest::send_file(
        //     $fileData,
        //     'file.xlsx',
        //     'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        // );
    }

    /**
     * Handle the export, for both the action button and the URL
      */
    public function handleXls(GridField $gridField, $request = null)
    {
        $items = $this->getItems($gridField);

        $this->setHeader($gridField, 'xls');

        $formater = new OldExcelDataFormatter();
        $fileData = $formater->convertDataObjectSet($items);

        return $fileData;
    }

    /**
     * Handle the export, for both the action button and the URL
      */
    public function handleCsv(GridField $gridField, $request = null)
    {
        $items = $this->getItems($gridField);

        $this->setHeader($gridField, 'csv');

        $formater = new CsvDataFormatter();
        $fileData = $formater->convertDataObjectSet($items);

        return $fileData;
    }

    protected function setHeader($gridField, $ext)
    {
        $do = singleton($gridField->getModelClass());

        Controller::curr()->getResponse()
            ->addHeader(
                "Content-Disposition",
                'attachment; filename="' .
                $do->i18n_plural_name() .
                '.' . $ext . '"');
    }

    protected function getItems(GridField $gridField)
    {
        $gridField->getConfig()->removeComponentsByType('GridFieldPaginator');

        $items = $gridField->getManipulatedList();

        foreach ($gridField->getConfig()->getComponents() as $component) {
            if ($component instanceof GridFieldFilterHeader || $component instanceof GridFieldSortableHeader) {
                $items = $component->getManipulatedData($gridField, $items);
            }
        }

        $arrayList = new ArrayList();

        foreach ($items->limit(null) as $item) {
            if (!$item->hasMethod('canView') || $item->canView()) {
                $arrayList->add($item);
            }
        }

        return $arrayList;
    }
}
