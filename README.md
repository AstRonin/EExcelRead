EExcelRead
====================

Read excel file and save to DB.

Based on PHPExcel [http://www.codeplex.com/PHPExcel](http://www.codeplex.com/PHPExcel) (version 1.7.7, 2012-05-19)

##Requirements:
* PHP 5.2 +
* [Yii](http://yiiframework.com) 1.1 or above
* Extension used model [ActiveRecord](http://www.yiiframework.com/doc/guide/1.1/en/database.ar)
* [PHPExcel](http://www.codeplex.com/PHPExcel) library 

##Installation
Copy EExcelRead.php to:

        /protected/extensions/phpexcel/ or else
        
Add PHPExcel lybrary to:

        /protected/extensions/phpexcel/Classes/ or else

##Usage
###Simple exemple:
```php
<?php
    $this->widget('ext.phpexcel.EExcelRead', array(
        'inputFileName'=>'example1.xls',
        'modelString'=>'User',
    ));
```
        In this case model's properties will be fill in the same order as in excel file.
        Be careful, if not use 'modelFields' order fields in file must be equal order fields in model.

####Very simple exemple:
#####If use this method [How to upload a file using a model](http://www.yiiframework.com/wiki/2/):

```php
<?php
    $this->widget('ext.phpexcel.EExcelRead', array(
        'model'=>$this,
        'modelPropertyIncludeFile'=>'field_upload_file'
    ));
```
OR 

```php
<?php
    $this->widget('ext.phpexcel.EExcelRead', array(
        'model'=>$this,
    ));
```

###Configuration:
* libPath - The path to the PHP excel lib

        default: ext.phpexcel.Classes.PHPExcel
        
* inputFileType - If empty then type will be found automatically

        suported types: 'Excel2007','Excel5','Excel2003XML','OOCalc','SYLK','Gnumeric'
        default: empty string

* inputFileName - Name of excel file
* skipRows - You can skip row(s). Supported types: boolean|int|array (not string).
    
        boolean = true - that skiped first row.
        int - that skiped one row. Int must be set exactly as in the file. First row = 1
        array - that skiped each row from array. array(1,5,69)

* modelString - Name of Model. String.
* model - The model instance. Object.

        Can set modelString as string or model as object or both.
        
    ```php
        ...
        'modelString' => 'MyModel',
        'model'=>$this, // new MyElseModel() if use another model
        ...
    ``` 

* modelPropertyIncludeFile - Extension will try find file in model property if upload file will be through model. See CUploadedFile

    ```php
        ...
        'model'=>$this,
        'modelPropertyIncludeFile'=>'field_upload_file'
        ...
    ```

* modelDefaultFields - Fill fields static data which not set in the file or replace existing

    ```php
        ...
        'modelDefaultFields'=>array(
           'user_id'=>10,
           'visible'=>true,
        ),
        ...
    ```

* modelFields - Array of model properties, sequence by DB's table.
        
        if array's value = 0|''|null|false, not action under this field

    ```php
        ...
        'modelFields'=>array(
            '',
            'name',
            'age',
            'comm',
        ),
        ...
    ```
        for update row in the DB and if excel has identification field(like ID) add field 'id'
        
    ```php
        ...
        'modelFields'=>array(
            'id',
            'name',
            ...
        ),
        ...
    ```

        If excel have fild from other model:
        
    ```php
        ...
        'modelFields'=>array(
            'name',
            ...
            'post'=>array(
                'class'=>'Post',
                'relationKey'=>'user_id',
                'defaultValues'=>array('visible'=>1),
            ),
            'address'=>array(
                'class'=>'Adress',
                'relationKey'=>'user_id',
            ),
            ...
        ),
        ...
    ```
            class - Name of relation model.
            relationKey - Surrogate key.
            defaultValues - works like as 'modelDefaultFields'.

* useRowAsModelFields - Use first (or else) row from the excel file as model fields

          Type: boolean | int
          If set TRUE selected first row. Integer must be set exactly as in the file!
          Attention! List of fields in the modelFields has a higher priority. If filled modelFields
          that useRowAsModelFields is ignored.
    
          default = false

* onBeforeSave - Call function before save. 
* onAfterSave - Call functon after save

          'MyMethod' | array($this,'MyMethod') | array('MyClass', 'MyMethod')
          
          Exemple: 'onBeforeSave'=>array($this,'uploadCallbackBeforeSave'),
        
##Change Log

Version 0.4 (21.09.2012)

* Add new property: modelPropertyIncludeFile, model
* Change work of property: modelString, now only as string
* Add automatic manipulation with file used CUploadedFile
* Add two methods of simple start the extension
* Bug fixing...





PS: Extension is written after using EExcelView (unashamedly taken the idea of callbacks).