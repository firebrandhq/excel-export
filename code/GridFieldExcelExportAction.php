<?php

/**
 * Gridfield component that can be added to a Gridfield to allow a user to export a single DataObject to Excel.
 *
 * Based of {@link GridFieldDeleteAction}.
 */
class GridFieldExcelExportAction implements GridField_ColumnProvider, GridField_ActionProvider {

    /**
     * The type of file we will be exporting
     * @var string
     */
    protected $exportType;


    /**
     * Instanciate a new GridFieldExcelExportAction
     * @param string $exportType The type of file we will be exporting. Defaults to 'xlsx', but 'csv' and 'xls' are also
     * acceptable.
     */
    public function __construct($exportType = 'xlsx') {
        $this->exportType = $exportType;
    }

    /**
     * Add a column at the end of the grid field if need be
     *
     * @param GridField $gridField
     * @param array $columns
     */
    public function augmentColumns($gridField, &$columns) {
        if(!in_array('Actions', $columns)) {
            $columns[] = 'Actions';
        }
    }

    /**
     * Return any special attributes that will be used for FormField::create_tag()
     *
     * @param GridField $gridField
     * @param DataObject $record
     * @param string $columnName
     * @return array
     */
    public function getColumnAttributes($gridField, $record, $columnName) {
        return array('class' => 'col-buttons');
    }

    /**
     * Add the title
     *
     * @param GridField $gridField
     * @param string $columnName
     * @return array
     */
    public function getColumnMetadata($gridField, $columnName) {
        if($columnName == 'Actions') {
            return array('title' => '');
        }
    }

    /**
     * Which columns are handled by this component
     *
     * @param GridField $gridField
     * @return array
     */
    public function getColumnsHandled($gridField) {
        return array('Actions');
    }

    /**
     * Which GridField actions are this component handling
     *
     * @param GridField $gridField
     * @return array
     */
    public function getActions($gridField) {
        return array('exportsingle');
    }

    /**
     * Return the button to show at the end of the row
     * @param GridField $gridField
     * @param DataObject $record
     * @param string $columnName
     * @return string - the HTML for the column
     */
    public function getColumnContent($gridField, $record, $columnName) {
        if(!$record->canView()) return;

        $field = GridField_FormAction::create($gridField, 'ExportSingle'.$record->ID, false,
                "exportsingle", array('RecordID' => $record->ID))
            ->addExtraClass('gridfield-button-export-single no-ajax')
            ->setAttribute('title', _t('firebrandhq.EXCELEXPORT', "Export"))
            ->setAttribute('data-icon', 'download-csv');
        return $field->Field();
    }

    /**
     * Handle the actions and apply any changes to the GridField
     *
     * @param GridField $gridField
     * @param string $actionName
     * @param mixed $arguments
     * @param array $data - form data
     * @return void
     */
    public function handleAction(GridField $gridField, $actionName, $arguments, $data) {
        if($actionName == 'exportsingle') {
            // Get the item
            $item = $gridField->getList()->byID($arguments['RecordID']);
            if(!$item) {
                return;
            }

            // Make sure th current user is authorised to view the item.
            if (!$item->canView()) {
                throw new ValidationException(
                    _t('firebrandhq.EXCELEXPORT', "Can not view record"),0);
            }

            // Build a filename
            $filename = $item->i18n_singular_name() . ' - ID ' . $item->ID;
            $title = $item->getTitle();
            if ($title) {
                $filename .= ' - ' . $title;
            }
            
            // Pick Converter
            switch ($this->exportType) {
                case 'xlsx':
                    $formater = new ExcelDataFormatter();
                    break;
                case 'xls':
                    $formater = new OldExcelDataFormatter();
                    break;
                case 'csv':
                    $formater = new CsvDataFormatter();
                    break;
                default:
                    user_error(
                        "GridFieldExcelExportAction expects \$exportType to be either 'xlsx', 'xls' or 'csv'. " .
                        "'{$this->exportType}' provided",
                        E_USER_ERROR
                    );
            }

            // Set the header that will cause the browser to download and save the file.
            $this->setHeader($gridField, $this->exportType, $filename);

            // Export our Data Object
            $fileData = $formater->convertDataObject($item);

            return $fileData;
        }
    }

    /**
     * Helper function for building the right header to get the file downloaded.
     * @param GridField $gridField
     * @param string $ext Extension the file should have
     * @param string $filename Optional of the file (without the extension). Defaults to the grid field object type.
     */
    protected function setHeader($gridField, $ext, $filename = '')
    {
        $do = singleton($gridField->getModelClass());
        if (!$filename) {
            $filename = $do->i18n_plural_name();
        }

        Controller::curr()->getResponse()
            ->addHeader(
                "Content-Disposition",
                'attachment; filename="' .
                $filename .
                '.' . $ext . '"'
            );
    }
}
