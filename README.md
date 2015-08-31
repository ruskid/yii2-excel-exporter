# Yii2 CSV Exporter
This will export Active Records into Excel file

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist ruskid/yii2-excel-exporter "*"
```

or add

```
"ruskid/yii2-excel-exporter": "*"
```

to the require section of your `composer.json` file.

Usage
-----

```php
$exporter = new CSVExporter;
$exporter->models = User::find()->all();
$exporter->values = [
    [
        'header' => Yii::t('app', 'Email'),
        'value' => function($object) {
            return $object->email;
        }
    ],
    [
        'header' => Yii::t('app', 'Nombre'),
        'value' => function($object) {
            return $object->nombre;
        }
    ],
    [
        'header' => Yii::t('app', 'Parent'),
        'value' => function($object) {
            return $object->parent->nombre;
        }
    ],

];

//Optional
$exporter->filename("reports-2015");
$exporter->startPoint('B3');

//Send file to browser
return $exporter->export();
```
- You can also add <b>styling</b>, see style and headerStyle parameters. 