# Read and write simple Excel and CSV files

[![Latest Version on Packagist](https://img.shields.io/packagist/v/spatie/simple-excel.svg?style=flat-square)](https://packagist.org/packages/spatie/simple-excel)
![GitHub Workflow Status](https://img.shields.io/github/workflow/status/spatie/simple-excel/run-tests?label=tests)
[![Total Downloads](https://img.shields.io/packagist/dt/spatie/simple-excel.svg?style=flat-square)](https://packagist.org/packages/spatie/simple-excel)

This package allows you to easily read and write simple Excel and CSV files. Behind the scenes generators are used to ensure low memory usage, even when working with large files.

Here's an example on how to read an Excel or CSV.

```php
SimpleExcelReader::create($pathToFile)->getRows()
   ->each(function(array $rowProperties) {
        // process the row
    });
```

If `$pathToFile` ends with `.csv` a CSV file is assumed. If it ends with `.xlsx`, an Excel file is assumed.

## Support us

Learn how to create a package like this one, by watching our premium video course:

[![Laravel Package training](https://spatie.be/github/package-training.jpg)](https://laravelpackage.training)

We invest a lot of resources into creating [best in class open source packages](https://spatie.be/open-source). You can support us by [buying one of our paid products](https://spatie.be/open-source/support-us).

We highly appreciate you sending us a postcard from your hometown, mentioning which of our package(s) you are using. You'll find our address on [our contact page](https://spatie.be/about-us). We publish all received postcards on [our virtual postcard wall](https://spatie.be/open-source/postcards).

## Installation

You can install the package via composer:

```bash
composer require spatie/simple-excel
```

## Usage

### Reading a CSV

Imagine you have a CSV with this content.

```csv
email,first_name
john@example.com,john
jane@example.com,jane
```

```php
// $rows is an instance of Illuminate\Support\LazyCollection
$rows = SimpleExcelReader::create($pathToCsv)->getRows();

$rows->each(function(array $rowProperties) {
   // in the first pass $rowProperties will contain
   // ['email' => 'john@example.com', 'first_name' => 'john']
});
```

#### Reading an Excel file

Reading an Excel file is identical to reading a CSV file. Just make sure that the path given to the `create` method of `SimpleExcelReader` ends with `xlsx`.

#### Working with LazyCollections

`getRows` will return an instance of [`Illuminate\Support\LazyCollection`](https://laravel.com/docs/master/collections#lazy-collections). This class is part of the Laravel framework. Behind the scenes generators are used, so memory usage will be low, even for large files.

You'll find a list of methods you can use on a `LazyCollection` [in the Laravel documentation](https://laravel.com/docs/master/collections#the-enumerable-contract).

Here's a quick, silly example where we only want to process rows that have a `first_name` that contains more than 5 characters.

```php
SimpleExcelReader::create($pathToCsv)->getRows()
    ->filter(function(array $rowProperties) {
       return strlen($rowProperties['first_name']) > 5
    })
    ->each(function(array $rowProperties) {
        // processing rows
    });
```

#### Reading a file without titles

If the file you are reading does not contain a title row, then you should use the `noHeaderRow()` method.

```php
// $rows is an instance of Illuminate\Support\LazyCollection
$rows = SimpleExcelReader::create($pathToCsv)
    ->noHeaderRow()
    ->getRows()
    ->each(function(array $rowProperties) {
       // in the first pass $rowProperties will contain
       // [0 => 'john@example', 1 => 'john']
});
```

#### Manually working with the reader object

Under the hood this package uses the [box/spout](https://github.com/box/spout) package. You can get to the underlying reader that implements `\Box\Spout\Reader\ReaderInterface` by calling the `getReader` method.

```php
$reader = SimpleExcelReader::create($pathToCsv)->getReader();
```

#### Offset & Limit

The `take` method allows you to specify a limit on how many rows are returned in the LazyCollection. The `skip` method allows you to define which row to start reading data from.

```php
// $rows is an instance of Illuminate\Support\LazyCollection
$rows = SimpleExcelReader::create($pathToCsv)
    ->skip(10)
    ->take(5)
    ->getRows();
```

### Writing files

Here's how you can write a CSV file:

```php
$writer = SimpleExcelWriter::create($pathToCsv)
     ->addRow([
        'first_name' => 'John',
        'last_name' => 'Doe',
    ])
    ->addRow([
        'first_name' => 'Jane',
        'last_name' => 'Doe',
    ]);
```

The file at `pathToCsv` will contain:

```csv
first_name,last_name
John,Doe
Jane,Doe
```

#### Writing an Excel file

Writing an Excel file is identical to writing a csv. Just make sure that the path given to the `create` method of `SimpleExcelWriter` ends with `xlsx`.

#### Streaming an Excel file to the browser

Instead of writing a file to disk, you can stream it directly to the browser.

```php
$writer = SimpleExcelWriter::streamDownload('your-export.xlsx')
     ->addRow([
        'first_name' => 'John',
        'last_name' => 'Doe',
    ])
    ->addRow([
        'first_name' => 'Jane',
        'last_name' => 'Doe',
    ])
    ->toBrowser();
```

### Writing multiple rows at once

You can use `addRows` instead of `addRow` to add multiple rows at once.

```php
$writer = SimpleExcelWriter::streamDownload('your-export.xlsx')
     ->addRows([
        [
            'first_name' => 'John',
            'last_name' => 'Doe',
        ],
        [
            'first_name' => 'Jane',
            'last_name' => 'Doe',
        ],
    ]);
```

#### Writing a file without titles

If the file you are writing should not have a title row added automatically, then you should use the `noHeaderRow()` method.

```php
$writer = SimpleExcelWriter::create($pathToCsv)
    ->noHeaderRow()
    ->addRow([
        'first_name' => 'Jane',
        'last_name' => 'Doe',
    ]);
});
```

This will output:

```csv
Jane,Doe
```

#### Adding layout

Under the hood this package uses the [box/spout](https://github.com/box/spout) package. That package contains a `StyleBuilder` that you can use to format rows. Styles can only be used on excel documents.

```php
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Common\Entity\Style\Color;

$style = (new StyleBuilder())
   ->setFontBold()
   ->setFontSize(15)
   ->setFontColor(Color::BLUE)
   ->setShouldWrapText()
   ->setBackgroundColor(Color::YELLOW)
   ->build();

$writer->addRow(['values, 'of', 'the', 'row'], $style)
```

For more information on styles head over to [the Spout docs](https://opensource.box.com/spout/docs/#styling).

#### Using an alternative delimiter

By default the `SimpleExcelReader` will assume that the delimiter is a `,`.

This is how you can use an alternative delimiter:

```php
SimpleExcelWriter::create($pathToCsv)->useDelimiter(';');
```

#### Getting the number of rows written

You can get the number of rows that are written. This number includes the automatically added header row.

```php
$writerWithAutomaticHeader = SimpleExcelWriter::create($this->pathToCsv)
    ->addRow([
        'first_name' => 'John',
        'last_name' => 'Doe',
    ]);

$writerWithoutAutomaticHeader->getNumberOfRows() // returns 2
```

#### Manually working with the writer object

Under the hood this package uses the [box/spout](https://github.com/box/spout) package. You can get to the underlying writer that implements `\Box\Spout\Reader\WriterInterface` by calling the `getWriter` method.

```php
$writer = SimpleExcelWriter::create($pathToCsv)->getWriter();
```

### Testing

``` bash
composer test
```

### Changelog

Please see [CHANGELOG](CHANGELOG.md) for more information on what has changed recently.

## Contributing

Please see [CONTRIBUTING](CONTRIBUTING.md) for details.

### Security

If you discover any security related issues, please email freek@spatie.be instead of using the issue tracker.

## Postcardware

You're free to use this package, but if it makes it to your production environment we highly appreciate you sending us a postcard from your hometown, mentioning which of our package(s) you are using.

Our address is: Spatie, Kruikstraat 22, 2018 Antwerp, Belgium.

We publish all received postcards [on our company website](https://spatie.be/en/opensource/postcards).

## Credits

- [Freek Van der Herten](https://github.com/freekmurze)
- [All Contributors](../../contributors)

## Alternatives

- [PhpSpreadsheet](https://phpspreadsheet.readthedocs.io/en/latest/)
- [laravel-excel](https://laravel-excel.com)

## License

The MIT License (MIT). Please see [License File](LICENSE.md) for more information.
