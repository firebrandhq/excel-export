<?php

/**
 * Enhanced GridField export button that allows the list to be exported to:
 *  * Excel 2007,
 *  * Excel 5,
 *  * CSV
 *
 * The button appears has a Split button exposing the 3 possible export format.
 *
 * @author Firebrand <hello@firebrand.nz>
 * @license MIT
 * @package silverstripe-excel-export
 */
class GridFieldExcelExportButton implements
    GridField_HTMLProvider,
    GridField_ActionProvider,
    GridField_URLHandler
{

    /**
     * Fragment to write the button to
     */
    protected $targetFragment;

    /**
     * Instanciate GridFieldExcelExportButton.
     * @param string $targetFragment
     */
    public function __construct($targetFragment = "before")
    {
        $this->targetFragment = $targetFragment;
    }

    /**
     * @inheritdoc
     *
     * Create the split button with all the export options.
     *
     * @param  GridField $gridField
     * @return array
     */
    public function getHTMLFragments($gridField)
    {
        // Set up the split button
        $splitButton = new SplitButton('Export', 'Export');
        $splitButton->setAttribute('data-icon', 'download-csv');

        // XLSX option
        $button = new GridField_FormAction(
            $gridField,
            'xlsxexport',
            _t('firebrandhq.EXCELEXPORT', 'Export to Excel (XLSX)'),
            'xlsxexport',
            null
        );
        $button->addExtraClass('no-ajax');
        $splitButton->push($button);

        // XLS option
        $button = new GridField_FormAction(
            $gridField,
            'xlsexport',
            _t('firebrandhq.EXCELEXPORT', 'Export to Excel (XLS)'),
            'xlsexport',
            null
        );
        $button->addExtraClass('no-ajax');
        $splitButton->push($button);

        // CSV option
        $button = new GridField_FormAction(
            $gridField,
            'csvexport',
            _t('firebrandhq.EXCELEXPORT', 'Export to CSV'),
            'csvexport',
            null
        );
        $button->addExtraClass('no-ajax');
        $splitButton->push($button);

        // Return the fragment
        return array(
            $this->targetFragment =>
                 $splitButton->Field()
        );
    }

    /**
     * @inheritdoc
     */
    public function getActions($gridField)
    {
        return array('xlsxexport', 'xlsexport', 'csvexport');
    }

    /**
     * @inheritdoc
     */
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
     * @inheritdoc
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
     * Action to export the GridField list to an Excel 2007 file.
     * @param  GridField $gridField
     * @param  SS_HTTPRequest    $request
     * @return string
     */
    public function handleXlsx(GridField $gridField, $request = null)
    {
        return $this->genericHandle('ExcelDataFormatter', 'xlsx', $gridField, $request);
    }

    /**
     * Action to export the GridField list to an Excel 5 file.
     * @param  GridField $gridField
     * @param  SS_HTTPRequest    $request
     * @return string
     */
    public function handleXls(GridField $gridField, $request = null)
    {
        return $this->genericHandle('OldExcelDataFormatter', 'xls', $gridField, $request);
    }

    /**
     * Action to export the GridField list to an CSV file.
     * @param  GridField $gridField
     * @param  SS_HTTPRequest    $request
     * @return string
     */
    public function handleCsv(GridField $gridField, $request = null)
    {
        return $this->genericHandle('CsvDataFormatter', 'csv', $gridField, $request);
    }

    /**
     * Generic Handle request that will return a Spread Sheet in the requested format
     * @param  string    $dataFormatterClass
     * @param  string    $ext
     * @param  GridField $gridField
     * @param  SS_HTTPRequest    $request
     * @return string
     */
    protected function genericHandle($dataFormatterClass, $ext, GridField $gridField, $request = null)
    {
        $items = $this->getItems($gridField);

        $this->setHeader($gridField, $ext);

        $formater = new $dataFormatterClass();
        $fileData = $formater->convertDataObjectSet($items);

        return $fileData;
    }

    /**
     * Set the HTTP header to force a download and set the filename.
     * @param GridField $gridField
     * @param string $ext Extension to use in the filename.
     */
    protected function setHeader($gridField, $ext)
    {
        $do = singleton($gridField->getModelClass());

        Controller::curr()->getResponse()
            ->addHeader(
                "Content-Disposition",
                'attachment; filename="' .
                $do->i18n_plural_name() .
                '.' . $ext . '"'
            );
    }

    /**
     * Helper function to extract the item list out of the GridField.
     * @param  GridField $gridField
     * @return SS_list
     */
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
