<?php

/**
 * EExcelRead
 * 
 * (PHP 5)
 * 
 * Read excel file and save to DB
 * Based on PHPExcel http://www.codeplex.com/PHPExcel | version 1.7.7, 2012-05-19
 * 
 * <b>Example</b>:
 * 
 * $this->widget('ext.phpexcel.EExcelRead', array(
 *           'inputFileName'=>'example1.xls',
 *           'skipRows'=>true,
 *           'modelString'=>'User',
 *           'onBeforeSave'=>array($this,'uploadCallbackBeforeSave'),
 *           'modelDefaultFields'=>array(
 *               'some_field'=>true,
 *               'some_field2'=>10,
 *           ),
 *           'modelFields'=>array(
 *               'name',
 *               'age',
 *               'comm',
 *               'post'=>array(
 *                   'class'=>'Post',
 *                   'relationKey'=>'user_id',
 *                   'defaultValues'=>array('visible'=>1),
 *              ),
 *              'address'=>array(
 *                   'class'=>'Adress',
 *                   'relationKey'=>'user_id',
 *              ),
 *           ),
 *       ));
 * 
 * If not set modelFields then fill all model properties.
 * 
 * @author Roman Shuplov <astronin@gmail.com>
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 * @version 0.3 Betta
 */
class EExcelRead extends CWidget {

    /**
     *
     * @var PHPExcel
     */
    public $objPHPExcel = null;

    /**
     * The path to the PHP excel lib
     *
     * @var string
     */
    public $libPath = 'ext.phpexcel.Classes.PHPExcel';

    /**
     * If empty then type will be found automatically
     *
     * <b>suported types</b>: 'Excel2007','Excel5','Excel2003XML','OOCalc','SYLK','Gnumeric'
     * 
     * <b>default</b>: empty
     * 
     * @var string 
     */
    public $inputFileType = '';
    public $inputFileName = null;

    /**
     * Sheet data from file as array
     *
     * @var arrya
     */
    public $sheetData = null;

    /**
     * Name of Model
     *
     * @var string
     */
    public $modelString = null;

    /**
     * Array of model properties, sequence by DB's table
     * 
     * If array's value = 0|''|null|false, not action under this field
     *
     * @var array
     */
    public $modelFields = array();

    /**
     * Fill fields static data which not set in the file or replace existing
     * 
     * For exemple
     * <pre>
     * array(
     *     'user_id'=>10,
     *     'visible'=>true,
     * );
     * </pre>
     * 
     * @var array
     */
    public $modelDefaultFields = array();

    /**
     * Use first or else row from the file as model fields
     * 
     * <b>Type</b>: boolean | int<br>
     * If set <b>true</b> selected first row. Int must be set exactly as in the file!
     * 
     * Attention! List of fields in the modelFields has a higher priority. If filled modelFields
     * that useRowAsModelFields is ignored.
     *
     * @var mixed
     */
    public $useRowAsModelFields = false;

    /**
     * Skip row.
     * 
     * <b>Types</b>: boolean|int|array (not string)
     * 
     * If you can set as <b>boolean</b> = true that skiped first row.<br>
     * If you can set as <b>int</b> that skiped one row. Int must be set exactly as in the file!<br>
     * If you can set as <b>arrya</b> that skiped each row from array.<br>
     *
     * @var mixed
     */
    public $skipRows = null;
    
    /**
     * Call function before save
     * 
     * 'MyMethod' | array($this,'MyMethod') | array('MyClass', 'MyMethod')
     *
     * @var mixed
     */
    public $onBeforeSave;
    
    /**
     * Call functon after save
     * 
     * 'MyMethod' | array($this,'MyMethod') | array('MyClass', 'MyMethod')
     *
     * @var mixed
     */
    public $onAfterSave;

    public function run() {
        if (!$this->createObjPHPExcel()) {
            Yii::log("Object PHPExcel not load", CLogger::LEVEL_ERROR, 'EExcelRead');
            return;
        }
        if (!$this->gettingSheetData()) {
            Yii::log("File does not have data", CLogger::LEVEL_INFO, 'EExcelRead');
            return;
        }

        /**
         * @todo remove code if not needed.
         */
        /*
          if ( count($sheetData) !== count($this->modelFields) ) {
          Yii::log('Count of user and document fields not equal');
          return false;
          }
         */

        if (is_object($this->modelString)) {
            $this->modelString = get_class($this->modelString);
        }

        $this->checkModelFieldsInFile();

        $this->iteration();

        parent::run();
    }

    private function checkModelFieldsInFile() {
        if ($this->useRowAsModelFields && !$this->modelFields) {
            if ($this->useRowAsModelFields === true) {
                $this->modelFields = $this->sheetData[0];
            }
            if (is_int($this->useRowAsModelFields)) {
                if ($this->useRowAsModelFields === 0)
                    $this->useRowAsModelFields = 1;
                $this->modelFields = $this->sheetData[($this->useRowAsModelFields - 1)];
            }
        }
    }

    private function iteration() {
        $sheetRowIterator = 0;
        foreach ($this->sheetData as $sheetRow) {

            if ($this->isSkipRow($sheetRowIterator)) {
                $sheetRowIterator++;
                continue;
            }

            /* @var $model CActiveRecord */
            $model = new $this->modelString;
            /* @var $relationModels CActiveRecord[] */
            $relationModels = array();
            $i = 0;
            if (!$this->modelFields) {
                $arrayModel = (array) $model;
                foreach ($arrayModel as $modelProperty => $v) {
                    if (!isset($sheetRow[$i]))
                        break;
                    $model->$modelProperty = $sheetRow[$i];
                    $i++;
                }
            } elseif ($this->modelFields && is_array($this->modelFields)) {
                foreach ($this->modelFields as $key => $modelProperty) {

                    // if count fields of file les count fields of config
                    if (!isset($sheetRow[$i]))
                        break;

                    if ($modelProperty) {
                        if (is_array($modelProperty)) {
                            $relationModels[] = $this->addRelationModel($key, $modelProperty, $sheetRow[$i]);
                        } else {
                            $model->$modelProperty = $sheetRow[$i];
                        }
                    }
                    $i++;
                }
            }
            if ($this->modelDefaultFields && is_array($this->modelDefaultFields))
                foreach ($this->modelDefaultFields as $keyForRelationModel => $valueForRelationModel) {
                    $model->$keyForRelationModel = $valueForRelationModel;
                }

            if ($model->id) {
                $model->setIsNewRecord(false);
            }else{
                $model->id = null;
            }
                
            if(is_callable($this->onBeforeSave))
                call_user_func_array($this->onBeforeSave, array($model, $relationModels, $sheetRow));
            
            if ($model->save()) {
                foreach ($relationModels as $a) {
                    if ($a)
                        if ($relationModel = $a[0]) {
                            if ($relationKey = $a[1])
                                $relationModel->$relationKey = $model->id;
                            if (!$relationModel->save()) {
                                Yii::log(print_r($relationModel->getErrors(), true), CLogger::LEVEL_WARNING, 'EExcelRead');
                            }
                        }
                }
                
                if(is_callable($this->onAfterSave))
                    call_user_func_array($this->onAfterSave, array($model, $relationModels, $sheetRow));
                
            } else {
                Yii::log(print_r($model->getErrors(), true), CLogger::LEVEL_WARNING, 'EExcelRead');
            }
            $sheetRowIterator++;
        }
    }

    /**
     * Check if need skip row(s)
     * 
     * @see EExcelRead::$skipRows
     * @param int $sheetRowIterator
     * @return boolean
     */
    private function isSkipRow($sheetRowIterator) {
        if ($this->skipRows !== null && $this->skipRows !== false && $this->useRowAsModelFields) {
            $this->skipRows = $this->useRowAsModelFields;
        }

        if ($this->skipRows !== null && $this->skipRows !== false) {
            if (is_array($this->skipRows)) {
                if (in_array(($sheetRowIterator + 1), $this->skipRows))
                    return true;
            }elseif ($this->skipRows === true) {
                if ($sheetRowIterator === 0)
                    return true;
            }elseif (is_int($this->skipRows)) {
                if ($this->skipRows === 0)
                    $this->skipRows = 1;
                if ($sheetRowIterator === ($this->skipRows - 1))
                    return true;
            }
        }
        return false;
    }

    /**
     * 
     * @return EExcelRead
     */
    private function createObjPHPExcel() {
        try {
            $lib = Yii::getPathOfAlias($this->libPath) . '.php';
            if (!file_exists($lib)) {
                Yii::log("PHP Excel lib not found($lib). Read disabled !", CLogger::LEVEL_WARNING, 'EExcelRead');
                return false;
            }
            spl_autoload_unregister(array('YiiBase', 'autoload'));
            Yii::import($this->libPath, true);

            if (!$this->inputFileType) {
                $this->inputFileType = PHPExcel_IOFactory::identify($this->inputFileName);
            }
            $objReader = PHPExcel_IOFactory::createReader($this->inputFileType);
            $objReader->setReadDataOnly(true);
            $this->objPHPExcel = $objReader->load($this->inputFileName);
            spl_autoload_register(array('YiiBase', 'autoload'));

            return true;
        } catch (Exception $e) {
            Yii::log($e->getTraceAsString(), CLogger::LEVEL_ERROR, 'EExcelRead');
        }
    }

    /**
     * Geting sheet data as array
     * 
     * @return int Count of array
     */
    private function gettingSheetData() {
        if (!$this->sheetData) {
            if ($this->objPHPExcel)
                $this->sheetData = $this->objPHPExcel->getActiveSheet()->toArray('', true, true, false);
        }
        return count($this->sheetData);
    }

    private function addRelationModel($relationModelProperty, $relationModelConfig, $sheetFieldData) {
        if (!isset($relationModelConfig['class'])) {
            Yii::log('Relation config has no key "class"', CLogger::LEVEL_ERROR, 'EExcelRead');
            return array();
        }
        if (!$relationModelConfig['class']) {
            Yii::log('Relation config has no value for "class"', CLogger::LEVEL_ERROR, 'EExcelRead');
            return array();
        }
        /* @var $relationModel CActiveRecord */
        try {
            $relationModel = new $relationModelConfig['class'];
        } catch (Exception $e) {
            $relationModel = null;
            Yii::log($e->getTraceAsString(), CLogger::LEVEL_ERROR, 'EExcelRead');
        }
        $relationModel->$relationModelProperty = $sheetFieldData;
        if (isset($relationModelConfig['defaultValues']) && $relationModelConfig['defaultValues'] && is_array($relationModelConfig['defaultValues'])) {
            foreach ($relationModelConfig['defaultValues'] as $key => $value) {
                $relationModel->$key = $value;
            }
        }
        if (isset($relationModelConfig['relationKey']) && $relationModelConfig['relationKey']) {
            $relationKey = $relationModelConfig['relationKey'];
        } else {
            $relationKey = '';
        }
        return array($relationModel, $relationKey);
    }

}