<?php
class SplitButton extends TabSet
{

    protected $tab;

    public function __construct($name, $title=null) {
        $args = func_get_args();
        $name = array_shift($args);

        if ($args) {
            $title = array_shift($args);
        }

        if (!$title) {
            $title = self::name_to_label($name);
        }

		$this->tab = new Tab(
			'SplitButtonTab',
			$title
		);

        parent::__construct($name, $this->tab);

        $this->addExtraClass('ss-ui-action-tabset action-menus ss-ui-button');

        if($args) foreach($args as $button) {
            $isValidArg = (is_object($button) && !($button instanceof FormField));
            if (!$isValidArg) user_error('SplitButton::__construct(): Parameter not a valid FormField instance', E_USER_ERROR);
            $this->tab->push($button);
        }

        Requirements::css(EXCELEXPORT_DIR . '/css/splitbutton.css');
    }

    public function fieldByName($name)
    {
        return $this->tab->fieldByName();
    }

    public function fieldPosition($field)
    {
        return $this->tab->fieldPosition($field);
    }

    public function getChildren()
    {
        return $this->tab->getChildren();
    }

    public function setChildren($children)
    {
        return $this->tab->setChildren($children);
    }

    public function push(FormField $field)
    {
        return $this->tab->push($field);
    }

    public function insertBefore($field, $insertBefore)
    {
        return $this->tab->insertBefore($field, $insertBefore);
    }

    public function insertAfter($field, $insertBefore)
    {
        return $this->tab->insertAfter($field, $insertBefore);
    }

    public function removeByName($fieldName, $dataFieldOnly = false)
    {
        return $this->tab->removeByName($fieldName, $dataFieldOnly = false);
    }

    public function replaceField($fieldName, $newField)
    {
        return $this->tab->replaceField($fieldName, $newField);
    }
}
