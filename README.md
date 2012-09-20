EExcelRead

(PHP 5)

Read excel file and save to DB
Based on PHPExcel http://www.codeplex.com/PHPExcel | version 1.7.7, 2012-05-19

<b>Example</b>:

$this->widget('ext.phpexcel.EExcelRead', array(
          'inputFileName'=>'example1.xls',
          'skipRows'=>true,
          'modelString'=>'User',
          'onBeforeSave'=>array($this,'uploadCallbackBeforeSave'),
          'modelDefaultFields'=>array(
              'some_field'=>true,
              'some_field2'=>10,
          ),
          'modelFields'=>array(
              'name',
              'age',
              'comm',
              'post'=>array(
                  'class'=>'Post',
                  'relationKey'=>'user_id',
                  'defaultValues'=>array('visible'=>1),
             ),
             'address'=>array(
                  'class'=>'Adress',
                  'relationKey'=>'user_id',
             ),
          ),
      ));

If not set modelFields then fill all model properties.