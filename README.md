<a href="https://spatie.be/github-ad-click/simple-excel">
<img
    style="width: 100%; max-width: 100%;" alt="spatie-simple-excel"
    src="https://github-ads.s3.eu-central-1.amazonaws.com/simple-excel.jpg?t=1" >
</a>

<p align="center">
    <a href="https://packagist.org/packages/spatie/simple-excel">
        <img alt="Packagist" src="https://img.shields.io/packagist/v/spatie/simple-excel.svg?style=for-the-badge&logo=packagist">
    </a>
    <a href="https://github.com/spatie/simple-excel/actions?query=workflow%3Arun-tests+branch%3Amain">
        <img alt="Tests Passing" src="https://img.shields.io/github/actions/workflow/status/spatie/simple-excel/run-tests.yml?style=for-the-badge&logo=github&label=tests">
    </a>
    <a href="https://packagist.org/packages/spatie/simple-excel">
        <img alt="Downloads" src="https://img.shields.io/packagist/dt/spatie/simple-excel.svg?style=for-the-badge" >
    </a>
</p>

----

## [Table of contents - Docs 📑:](#table-of-contents)

## Support us


We invest a lot of resources into creating [best in class open source packages](https://spatie.be/open-source). You can support us by [buying one of our paid products](https://spatie.be/open-source/support-us).

We highly appreciate you sending us a postcard from your hometown, mentioning which of our package(s) you are using. You'll find our address on [our contact page](https://spatie.be/about-us). We publish all received postcards on [our virtual postcard wall](https://spatie.be/open-source/postcards).


<hr style="background-color: #1bb5d0">

# Spatie Simple Excel

This package allows you to easily read and write simple Excel and CSV files.
Behind the scenes generators are used to ensure low memory usage, even when working with large files.

----

- :fire: Great uses for **generators**
- :brain: **Low memory usage** working with **large files**
- :computer: **Read** simple **Excel** files
- :computer: **Write** simple **Excel** files
- :computer: **Read** simple **CSV** files
- :computer: **Write** simple **CSV** files
- :zap: Useful for **import** _(*)_
- :zap: Useful for **export** _(*)_
- :book: **Simplified** implementation
- :muscle: Great features with **few lines**

[*] _Example of implementation._

Table of contents
=================

- [Installation](#installation-)
- [Usage](#usage-)
  - [Reading a CSV](#reading-a-csv-)
    - [Reading an Excel file](#reading-an-excel-file-)
    - [Working with LazyCollections](#working-with-lazycollections-)
    - [Reading a file without headers](#reading-a-file-without-headers-)
    - [Manually setting the headers](#manually-setting-the-headers-)
  - [Working with multiple sheet documents](#working-with-multiple-sheet-documents-)
    - [Retrieving header row values](#retrieving-header-row-values-)
    - [Dealing with headers that are not on the first line](#dealing-with-headers-that-are-not-on-the-first-line-)
    - [Trimming headers](#trimming-headers-)
    - [Convert headers to snake_case](#convert-headers-to-snake_case-)
    - [Manually formatting headers](#manually-formatting-headers-)
    - [Manually working with the reader object](#manually-working-with-the-reader-object-)
    - [Limiting the result set](#limiting-the-result-set-)
  - [Writing files](#writing-files-)
    - [Manually set the header from array](#manually-set-the-header-from-array-)
    - [Writing an Excel file](#writing-an-excel-file-)
    - [Streaming an Excel file to the browser](#streaming-an-excel-file-to-the-browser-)
  - [Writing multiple rows at once](#writing-multiple-rows-at-once-)
    - [Writing a file without titles](#writing-a-file-without-titles-)
    - [Adding layout](#adding-layout-)
    - [Setting column widths and row heights](#setting-column-widths-and-row-heights-)
    - [Creating an additional sheets](#creating-an-additional-sheets-)
    - [Using an alternative delimiter](#using-an-alternative-delimiter-)
    - [Getting the number of rows written](#getting-the-number-of-rows-written-)
    - [Disable BOM](#disable-bom-)
    - [Manually working with the writer object](#manually-working-with-the-writer-object-)
- [Testing](#testing-)
- [Changelog](#changelog-)
- [Contributing](#contributing-)
- [Security](#security-)
- [Postcardware](#postcardware-)
- [Credits](#credits-)
- [Alternatives](#alternatives-)
- [License](#license-)

<hr style="background-color: #1bb5d0">

## Installation [^](#table-of-contents)

You can install the package via composer:

```bash
composer require spatie/simple-excel
```

## Usage [^](#table-of-contents)

Here's an example on how to read an Excel or CSV.

```php
use Spatie\SimpleExcel\SimpleExcelReader;

SimpleExcelReader::create($pathToFile)->getRows()
   ->each(function(array $rowProperties) {
        // process the row
    });
```

If `$pathToFile` ends with `.csv` a CSV file is assumed. If it ends with `.xlsx`, an Excel file is assumed.

### Reading a CSV [^](#table-of-contents)

Imagine you have a CSV with this content.

```csv
email,first_name
john@example.com,john
jane@example.com,jane
```

```php
use Spatie\SimpleExcel\SimpleExcelReader;

// $rows is an instance of Illuminate\Support\LazyCollection
$rows = SimpleExcelReader::create($pathToCsv)->getRows();

$rows->each(function(array $rowProperties) {
   // in the first pass $rowProperties will contain
   // ['email' => 'john@example.com', 'first_name' => 'john']
});
```

#### Reading an Excel file [^](#table-of-contents)

Reading an Excel file is identical to reading a CSV file. Just make sure that the path given to the `create` method of `SimpleExcelReader` ends with `xlsx`.

#### Working with LazyCollections [^](#table-of-contents)

`getRows` will return an instance of [`Illuminate\Support\LazyCollection`](https://laravel.com/docs/master/collections#lazy-collections). This class is part of the Laravel framework. Behind the scenes generators are used, so memory usage will be low, even for large files.

You'll find a list of methods you can use on a `LazyCollection` [in the Laravel documentation](https://laravel.com/docs/master/collections#the-enumerable-contract).

Here's a quick, silly example where we only want to process rows that have a `first_name` that contains more than 5 characters.

```php
SimpleExcelReader::create($pathToCsv)->getRows()
    ->filter(function(array $rowProperties) {
       return strlen($rowProperties['first_name']) > 5;
    })
    ->each(function(array $rowProperties) {
        // processing rows
    });
```

#### Reading a file without headers [^](#table-of-contents)

If the file you are reading does not contain a header row, then you should use the `noHeaderRow()` method.

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

#### Manually setting the headers [^](#table-of-contents)

If you would like to use a specific array of values for the headers, you can use the `useHeaders()` method.

```php
// $rows is an instance of Illuminate\Support\LazyCollection
$rows = SimpleExcelReader::create($pathToCsv)
    ->useHeaders(['email_address', 'given_name'])
    ->getRows()
    ->each(function(array $rowProperties) {
       // in the first pass $rowProperties will contain
       // ['email_address' => 'john@example', 'given_name' => 'john']
});
```

If your file already contains a header row, it will be ignored and replaced with your custom headers.

If your file does not contain a header row, you should also use `noHeaderRow()`, and your headers will be used instead of numeric keys, as above.

### Working with multiple sheet documents [^](#table-of-contents)

Excel files can include multiple spreadsheets. You can select the sheet you want to use with the `fromSheet()` method to select by index.

```php
$rows = SimpleExcelReader::create($pathToXlsx)
    ->fromSheet(3)
    ->getRows();
```

With multiple spreadsheets, you can too select the sheet you want to use with the `fromSheetName()` method to select by name.

```php
$rows = SimpleExcelReader::create($pathToXlsx)
    ->fromSheetName("sheet1")
    ->getRows();
```

#### Retrieving header row values [^](#table-of-contents)

If you would like to retrieve the header row as an array, you can use the `getHeaders()` method.

If you have used `useHeaders()` to set custom headers, these will be returned instead of the actual headers in the file. To get the original headers from the file, use `getOriginalHeaders()`.

```php
$headers = SimpleExcelReader::create($pathToCsv)->getHeaders();

// $headers will contain
// [ 'email', 'first_name' ]
```

#### Dealing with headers that are not on the first line [^](#table-of-contents)

If your file has headers that are not on the first line, you can use the `headerOnRow()` method
to indicate the line at which the headers are present. Any data above this line
will be discarded from the result.

`headerOnRow` accepts the line number as an argument, starting at 0. Blank lines are not counted.

Since blank lines will not be counted, this method is mostly useful for files
that include formatting above the actual dataset, which can be the case with Excel files.

```csv
This is my data sheet
See worksheet 1 for the data, worksheet 2 for the graphs.



email , firstname
john@example.com,john
jane@example.com,jane
```

```php
// $rows is an instance of Illuminate\Support\LazyCollection
$rows = SimpleExcelReader::create($pathToCsv)
    ->trimHeaderRow()
    ->headerOnRow(3)
    ->getRows()
    ->each(function(array $rowProperties) {
       // in the first pass $rowProperties will contain
       // ['email' => 'john@example', 'first_name' => 'john']
});
```

#### Trimming headers [^](#table-of-contents)

If the file you are reading contains a title row, but you need to trim additional characters on the title values, then you should use the `trimHeaderRow()` method.
This functionality mimics the `trim` method, and the default characters it trims, matches that function.

Imagine you have a csv file with this content.

```csv
email , first_name
john@example.com,john
jane@example.com,jane
```

```php
// $rows is an instance of Illuminate\Support\LazyCollection
$rows = SimpleExcelReader::create($pathToCsv)
    ->trimHeaderRow()
    ->getRows()
    ->each(function(array $rowProperties) {
       // in the first pass $rowProperties will contain
       // ['email' => 'john@example', 'first_name' => 'john']
});
```

`trimHeaderRow()` additionally accepts a param to specify what characters to trim. This param can utilize the same functionality allowed by the trim function's `$characters` definition including a range of characters.

#### Convert headers to snake_case [^](#table-of-contents)

If you would like all the headers to be converted to snake_case, use the the `headersToSnakeCase()` method.

```csv
Email,First Name,Last Name
john@example.com,john,doe
mary-jane@example.com,mary jane,doe
```

```php
$rows = SimpleExcelReader::create($pathToCsv)
    ->headersToSnakeCase()
    ->getRows()
    ->each(function(array $rowProperties) {
        // rowProperties converted to snake_case
        // ['email' => 'john@example', 'first_name' => 'John', 'last_name' => 'doe']
    });
```

#### Manually formatting headers [^](#table-of-contents)

You can use a custom formatter to change the headers using the `formatHeadersUsing` method and passing a closure.

```csv
email,first_name,last_name
john@example.com,john,doe
mary-jane@example.com,mary jane,doe
```

```php
$rows = SimpleExcelReader::create($pathToCsv)
    ->formatHeadersUsing(fn($header) => "{$header}_simple_excel")
    ->getRows()
    ->each(function(array $rowProperties) {
        // ['email_simple_excel' => 'john@example', 'first_name_simple_excel' => 'John', 'last_name_simple_excel' => 'doe']
    });
```

#### Manually working with the reader object [^](#table-of-contents)

Under the hood this package uses the [box/spout](https://github.com/openspout/openspout) package. You can get to the underlying reader that implements `\OpenSpout\Reader\ReaderInterface` by calling the `getReader` method.

```php
$reader = SimpleExcelReader::create($pathToCsv)->getReader();
```

#### Limiting the result set [^](#table-of-contents)

The `take` method allows you to specify a limit on how many rows should be returned.

```php
// $rows is an instance of Illuminate\Support\LazyCollection
$rows = SimpleExcelReader::create($pathToCsv)
    ->take(5)
    ->getRows();
```

The `skip` method allows you to define which row to start reading data from. In this example we get rows 11 to 16.


```php
$rows = SimpleExcelReader::create($pathToCsv)
    ->skip(10)
    ->take(5)
    ->getRows();
```

### Writing files [^](#table-of-contents)

Here's how you can write a CSV file:

```php
use Spatie\SimpleExcel\SimpleExcelWriter;

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

#### Manually set the header from array [^](#table-of-contents)

Instead of automatically let the package dedecting a header row, you can set it manually.

```php
use Spatie\SimpleExcel\SimpleExcelWriter;

$writer = SimpleExcelWriter::create($pathToCsv)
    ->addHeader(['first_name', 'last_name'])
    ->addRow(['John', 'Doe'])
    ->addRow(['Jane', 'Doe'])
```

#### Writing an Excel file [^](#table-of-contents)

Writing an Excel file is identical to writing a csv. Just make sure that the path given to the `create` method of `SimpleExcelWriter` ends with `xlsx`.


#### Streaming an Excel file to the browser [^](#table-of-contents)

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

Make sure to call `flush()` if you're sending large streams to the browser

```php
$writer = SimpleExcelWriter::streamDownload('your-export.xlsx');

foreach (range(1, 10_000) as $i) {
    $writer->addRow([
        'first_name' => 'John',
        'last_name' => 'Doe',
    ]);

    if ($i % 1000 === 0) {
        flush(); // Flush the buffer every 1000 rows
    }
}

$writer->toBrowser();
```


### Writing multiple rows at once [^](#table-of-contents)

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

#### Writing a file without titles [^](#table-of-contents)

If the file you are writing should not have a title row added automatically, then you should use the `noHeaderRow()` method.

```php
$writer = SimpleExcelWriter::create($pathToCsv)
    ->noHeaderRow()
    ->addRow([
        'first_name' => 'Jane',
        'last_name' => 'Doe',
    ]);
```

This will output:

```csv
Jane,Doe
```

#### Adding layout [^](#table-of-contents)

Under the hood this package uses the [openspout/openspout](https://github.com/openspout/openspout) package. That package contains a `Style` builder that you can use to format rows. Styles can only be used on excel documents.

```php
use OpenSpout\Common\Entity\Style\Color;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Entity\Style\Border;
use OpenSpout\Common\Entity\Style\BorderPart;

/* Create a border around a cell */
$border = new Border(
        new BorderPart(Border::BOTTOM, Color::LIGHT_BLUE, Border::WIDTH_THIN, Border::STYLE_SOLID),
        new BorderPart(Border::LEFT, Color::LIGHT_BLUE, Border::WIDTH_THIN, Border::STYLE_SOLID),
        new BorderPart(Border::RIGHT, Color::LIGHT_BLUE, Border::WIDTH_THIN, Border::STYLE_SOLID),
        new BorderPart(Border::TOP, Color::LIGHT_BLUE, Border::WIDTH_THIN, Border::STYLE_SOLID)
    );

$style = (new Style())
   ->setFontBold()
   ->setFontSize(15)
   ->setFontColor(Color::BLUE)
   ->setShouldWrapText()
   ->setBackgroundColor(Color::YELLOW)
   ->setBorder($border);

$writer->addRow(['values', 'of', 'the', 'row'], $style);
```
To style your HeaderRow simply call the `setHeaderStyle($style)` Method.

```php
$writer->setHeaderStyle($style);
```

For more information on styles head over to [the Spout docs](https://github.com/openspout/openspout/tree/4.x/docs).

#### Setting column widths and row heights [^](#table-of-contents)

By accessing the underlying OpenSpout Writer you can set default column widths and row heights and change the width of specific columns.

```php
SimpleExcelWriter::create(
    file: 'document.xlsx',
    configureWriter: function ($writer) {
        $options = $writer->getOptions;
        $options->DEFAULT_COLUMN_WIDTH=25; // set default width
        $options->DEFAULT_ROW_HEIGHT=15; // set default height
        // set columns 1, 3 and 8 to width 40
        $options->setColumnWidth(40, 1, 3, 8);
        // set columns 9 through 12 to width 10
        $options->setColumnWidthForRange(10, 9, 12);
    }
)
```

#### Creating an additional sheets [^](#table-of-contents)

By default, the writer will write to the first sheet. If you want to write to an additional sheet, you can use the `addNewSheetAndMakeItCurrent` method.

```php
$writer = SimpleExcelWriter::create($pathToXlsx);

Posts::all()->each(function (Post $post) use ($writer) {
    $writer->nameCurrentSheet($post->title);

    $post->comments->each(function (Comment $comment) use ($writer) {
        $writer->addRow([
            'comment' => $comment->comment,
            'author' => $comment->author,
        ]);
    });

    if(!$post->is($posts->last())) {
        $writer->addNewSheetAndMakeItCurrent();
    }
});
```

#### Using an alternative delimiter [^](#table-of-contents)

By default the `SimpleExcelReader` will assume that the delimiter is a `,`.

This is how you can use an alternative delimiter:

```php
SimpleExcelWriter::create(file: $pathToCsv, delimiter: ';');
```

#### Getting the number of rows written [^](#table-of-contents)

You can get the number of rows that are written. This number includes the automatically added header row.

```php
$writerWithAutomaticHeader = SimpleExcelWriter::create($this->pathToCsv)
    ->addRow([
        'first_name' => 'John',
        'last_name' => 'Doe',
    ]);

$writerWithAutomaticHeader->getNumberOfRows(); // returns 2
```

#### Disable BOM [^](#table-of-contents)

You can also disable adding a BOM to the start of the file. BOM must be disabled on create and cannot be disabled after creation of the writer.

A BOM, or byte order mark, indicates a number of things for the file being written including the file being unicode as well as it's UTF encoding type.

```php
SimpleExcelWriter::createWithoutBom($this->pathToCsv, $type);
```

Additional information about BOM can be found [here](https://en.wikipedia.org/wiki/Byte_order_mark).

#### Manually working with the writer object [^](#table-of-contents)

Under the hood this package uses the [openspout/openspout](https://github.com/openspout/openspout) package. You can get to the underlying writer that implements `\OpenSpout\Reader\WriterInterface` by calling the `getWriter` method.

```php
$writer = SimpleExcelWriter::create($pathToCsv)->getWriter();
```

## Testing [^](#table-of-contents)

``` bash
composer test
```

## Changelog [^](#table-of-contents)

Please see [CHANGELOG](CHANGELOG.md) for more information on what has changed recently.

## Contributing [^](#table-of-contents)

Please see [CONTRIBUTING](https://github.com/spatie/.github/blob/main/CONTRIBUTING.md) for details.

## Security [^](#table-of-contents)

If you've found a bug regarding security please mail [security@spatie.be](mailto:security@spatie.be) instead of using the issue tracker.

## Postcardware [^](#table-of-contents)

You're free to use this package, but if it makes it to your production environment we highly appreciate you sending us a postcard from your hometown, mentioning which of our package(s) you are using.

Our address is: Spatie, Kruikstraat 22, 2018 Antwerp, Belgium.

We publish all received postcards [on our company website](https://spatie.be/en/opensource/postcards).

## Credits [^](#table-of-contents)

- [Freek Van der Herten](https://github.com/freekmurze)
- [All Contributors](../../contributors)

## Alternatives [^](#table-of-contents)

- [PhpSpreadsheet](https://phpspreadsheet.readthedocs.io/en/latest/)
- [laravel-excel](https://laravel-excel.com)

## License [^](#table-of-contents)

The MIT License (MIT). Please see [License File](LICENSE.md) for more information.

----

<a href="https://spatie.be/github-ad-click/simple-excel">
<img
    style="width: 100%; max-width: 100%;" alt="spatie-simple-excel"
    src="https://github-ads.s3.eu-central-1.amazonaws.com/simple-excel.jpg?t=1" >
</a>

<p align="center">
    <a href="https://packagist.org/packages/spatie/simple-excel">
        <img alt="Packagist" src="https://img.shields.io/packagist/v/spatie/simple-excel.svg?style=for-the-badge&logo=packagist">
    </a>
    <a href="https://github.com/spatie/simple-excel/actions?query=workflow%3Arun-tests+branch%3Amain">
        <img alt="Tests Passing" src="https://img.shields.io/github/actions/workflow/status/spatie/simple-excel/run-tests.yml?style=for-the-badge&logo=github&label=tests">
    </a>
    <a href="https://packagist.org/packages/spatie/simple-excel">
        <img alt="Downloads" src="https://img.shields.io/packagist/dt/spatie/simple-excel.svg?style=for-the-badge" >
    </a>
</p>


----

## [Table of contents - Docs 📑:](#table-of-contents)

## Support us


We invest a lot of resources into creating [best in class open source packages](https://spatie.be/open-source). You can support us by [buying one of our paid products](https://spatie.be/open-source/support-us).

We highly appreciate you sending us a postcard from your hometown, mentioning which of our package(s) you are using. You'll find our address on [our contact page](https://spatie.be/about-us). We publish all received postcards on [our virtual postcard wall](https://spatie.be/open-source/postcards).

----
