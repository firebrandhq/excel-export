<?php

/**
 * Simple extension of {@link TabSet} for displaying a split button. This is
 * based off the action-menu logic used by SilverStripe for the _More options_
 * link on the SiteTree edit window.
 *
 * @author Firebrand <hello@firebrand.nz>
 * @license MIT
 * @package silverstripe-excel-export
 */
class SplitButton extends TabSet
{

    /**
     * Underlying container to witch the buttons are goign to be added.
     * @var Tab
     */
    protected $tab;

    /**
     * Create a new instance of SplitButton
     * @param string $name  Form field name
     * @param string $title Title that will be displayed on the split button.
     * if not provided, the title will be guess from the `$name`.
     */
    public function __construct($name, $title=null) {
        $args = func_get_args();
        $name = array_shift($args);

        if ($args) {
            $title = array_shift($args);
        }

        // Guess the title if none provided
        if (!$title) {
            $title = self::name_to_label($name);
        }

        // Instanciate our undelying tab container
        $this->tab = new Tab(
            'SplitButtonTab',
            $title
        );

        //Call the parent consturctor
        parent::__construct($name, $this->tab);

        // Add the same class as the _more options_ link so we can piggy back
        // off that logic.
        $this->addExtraClass('ss-ui-action-tabset action-menus ss-ui-button');

        // Add any provided button.
        if ($args) {
            foreach ($args as $button) {
                // Make sure we only add Form Fields to our tab.
                $isValidArg =
                    (is_object($button) &&
                    !($button instanceof FormField));
                if (!$isValidArg) {
                    user_error(
                        'SplitButton::__construct(): Parameter not a valid FormField instance',
                        E_USER_ERROR
                    );
                }

                $this->tab->push($button);
            }
        }

        // Define a custom spread sheet so we can style our button.
        Requirements::css(EXCELEXPORT_DIR . '/css/splitbutton.css');
    }

    /**
     * @inheritdoc
     */
    public function fieldByName($name)
    {
        return $this->tab->fieldByName();
    }

    /**
     * @inheritdoc
     */
    public function fieldPosition($field)
    {
        return $this->tab->fieldPosition($field);
    }

    /**
     * @inheritdoc
     */
    public function getChildren()
    {
        return $this->tab->getChildren();
    }

    /**
     * @inheritdoc
     */
    public function setChildren($children)
    {
        return $this->tab->setChildren($children);
    }

    /**
     * @inheritdoc
     */
    public function push(FormField $field)
    {
        return $this->tab->push($field);
    }

    /**
     * @inheritdoc
     */
    public function insertBefore($field, $insertBefore)
    {
        return $this->tab->insertBefore($field, $insertBefore);
    }

    /**
     * @inheritdoc
     */
    public function insertAfter($field, $insertBefore)
    {
        return $this->tab->insertAfter($field, $insertBefore);
    }

    /**
     * @inheritdoc
     */
    public function removeByName($fieldName, $dataFieldOnly = false)
    {
        return $this->tab->removeByName($fieldName, $dataFieldOnly = false);
    }

    /**
     * @inheritdoc
     */
    public function replaceField($fieldName, $newField)
    {
        return $this->tab->replaceField($fieldName, $newField);
    }
}
