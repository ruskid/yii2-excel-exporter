# Yii2 CSV Exporter
This will export Active Records into Excel file

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist ruskid/yii2-excel-exporter "dev-master"
```

or add

```
"ruskid/yii2-excel-exporter": "dev-master"
```

to the require section of your `composer.json` file.

Usage
-----

```php
 $exporter = new CSVExporter;
//Add headers A1-A3
$exporter->setCellValue('A1', 'Subcategoria');
$exporter->setCellValue('A2', 'Propietario');
$exporter->setCellValue('A3', 'Escenario');
//Apply header styles for A1-A3
for ($i = 1; $i <= 3; $i++) {
    $exporter->setCellStyleFromArray('A' . $i, $exporter->headerStyle);
}
//Set values to B1-B3
$exporter->setCellValue('B1', $model->name);
$exporter->setCellValue('B2', $model->propietario);
$exporter->setCellValue('B3', $model->escenario);

//Set server data. start from line A11. will return next free cell index
$nextCellIndex = $exporter->setCellData('A11', $model->servers, [
    [
        'header' => Yii::t('app', 'HOSTNAME'),
        'value' => function($array) {
            return $array['HOSTNAME'];
        }
    ],
    [
        'header' => Yii::t('app', 'DESC_CATALOGO'),
        'value' => function($array) {
            return $array['DESC_CATALOGO'];
        }
    ],
]);

//set users from next cell free cell index
$users = \app\models\User::find()->all();
$exporter->setCellData($nextCellIndex, $users, [
    [
        'header' => Yii::t('app', 'Login'),
        'value' => function($object) {
            return $object->username;
        }
    ],
]);

$exporter->filename = 'test';
return $exporter->export();
```
- You can also add <b>styling</b>, see style and headerStyle parameters. 