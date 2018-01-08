# Excel to files

`fdmsantos/exceltofiles` is a simple library to create files by excel row based on a template.

## How To Use

We need Give a Template File and The Excel File.

Template File Example:

```txt
Hello my name is [name].
I'm from [country] and have [age] years old!
````

Excel File Example:

| A      | B         | C  |
| -------|-----------| ---|
| Fabio  | Portugal  | 31 |
| John   | Spain     | 23 |
| Peter  | England   | 65 |

Code Example:

```php
require_once('vendor/autoload.php');

use ExcelToFiles\ExcelToFiles;

$exceltoFiles = new ExcelToFiles([
  'template' => 'template.txt',
  'excel'    => 'excel.xls',
  'mapping'  => [
    '[name]' => 'A'
    '[country]' => 'B'
    '[age]' => 'C'
]);

$exceltoFiles->generate();
```
The result from code above will be :

File 1:
```txt
Hello my name is Fabio.
I'm from Portugal and have 31 years old!
````

File 2:
```txt
Hello my name is John.
I'm from Spain and have 23 years old!
````

File 3:
```txt
Hello my name is Peter.
I'm from England and have 65 years old!
````

Optional Params:
```php
require_once('vendor/autoload.php');

use ExcelToFiles\ExcelToFiles;

$exceltoFiles = new ExcelToFiles([
  'template' => '{template.txt}',
  'excel'    => '{excel.xls}',
  'excludeRows' => [1,2,13], // To exclude Rows. This example wil exclude row 1, 2 and 13
  'filesname' => 'person_{A}.txt', // To define filename. The name can depends from excel row. For this it's necessary use {column}.
  'outputdir' => 'src/', // To choose files path. The Default is current path.
  'mapping'  => [
    '[name]' => 'A'
    '[country]' => 'B'
    '[age]' => 'C'
]);
````

Closures:

If the template variables have same logical or depends from two or more columns, we can use closures.

```php
// One Clousure for Variable
$exceltoFiles->mapWithClosure('[name]',function($columns) {
	return $columns['A'].' => '.$columns['B'];
});

$exceltoFiles->generate();
````
