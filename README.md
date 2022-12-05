
[<img src="https://github-ads.s3.eu-central-1.amazonaws.com/support-ukraine.svg?t=1" />](https://supportukrainenow.org)

# Read and write simple Excel and CSV files

[![Latest Version on Packagist](https://img.shields.io/packagist/v/spatie/simple-excel.svg?style=flat-square)](https://packagist.org/packages/spatie/simple-excel)
![GitHub Workflow Status](https://img.shields.io/github/workflow/status/spatie/simple-excel/run-tests?label=tests)
[![Total Downloads](https://img.shields.io/packagist/dt/spatie/simple-excel.svg?style=flat-square)](https://packagist.org/packages/spatie/simple-excel)

This package allows you to easily read and write simple Excel and CSV files. Behind the scenes generators are used to ensure low memory usage, even when working with large files.

Here's an example on how to read an Excel or CSV.

```php
use Spatie\SimpleExcel\SimpleExcelReader;

SimpleExcelReader::create($pathToFile)->getRows()
   ->each(function(array $rowProperties) {
        // process the row
    });
```

If `$pathToFile` ends with `.csv` a CSV file is assumed. If it ends with `.xlsx`, an Excel file is assumed.

## Support us

[<img src="https://github-ads.s3.eu-central-1.amazonaws.com/simple-excel.jpg?t=1" width="419px" />](https://spatie.be/github-ad-click/simple-excel)

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
use Spatie\SimpleExcel\SimpleExcelReader;

// $rows is an instance of Illuminate\Support\LazyCollection
$rows = SimpleExcelReader::create($pathToCsv)->getRows();

$rows->each(function(array $rowProperties) {
   // in the first pass $rowProperties will contain
   // ['email' => 'john@example.com', 'first_name' => 'john']
});
```

#### Reading an Excel file

Reading an Excel file is identical to reading a CSV file. Just make sure that the path given to the `create` method of `SimpleExcelReader` ends with `xlsx`.

#### Manually setting the file type

You can pass the file type to the `create` method of `SimpleExcelReader` as the second, optional argument:

```php
SimpleExcelReader::create($pathToFile, 'csv');
```

#### Working with LazyCollections

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

#### Reading a file without headers

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

#### Manually setting the headers

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

### Working with multiple sheet documents

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

#### Retrieving header row values

If you would like to retrieve the header row as an array, you can use the `getHeaders()` method.

If you have used `useHeaders()` to set custom headers, these will be returned instead of the actual headers in the file. To get the original headers from the file, use `getOriginalHeaders()`.

```php
$headers = SimpleExcelReader::create($pathToCsv)->getHeaders();

// $headers will contain
// [ 'email', 'first_name' ]
```

#### Dealing with headers that are not on the first line

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

#### Trimming headers

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

#### Convert headers to snake_case

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

#### Manually formatting headers

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

#### Manually working with the reader object

Under the hood this package uses the [box/spout](https://github.com/openspout/openspout) package. You can get to the underlying reader that implements `\OpenSpout\Reader\ReaderInterface` by calling the `getReader` method.

```php
$reader = SimpleExcelReader::create($pathToCsv)->getReader();
```

#### Limiting the result set

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

### Writing files

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

#### Manually set the header from array

Instead of automatically let the package dedecting a header row, you can set it manually.

```php
use Spatie\SimpleExcel\SimpleExcelWriter;

$writer = SimpleExcelWriter::create($pathToCsv)
    ->addHeader(['first_name', 'last_name'])
    ->addRow(['John', 'Doe'])
    ->addRow(['Jane', 'Doe'])
```

#### Writing an Excel file

Writing an Excel file is identical to writing a csv. Just make sure that the path given to the `create` method of `SimpleExcelWriter` ends with `xlsx`.

#### Manually setting the file type

You can pass the file type to the `create` method of `SimpleExcelWriter` as the second, optional argument:

```php
SimpleExcelWriter::create('php://output', 'csv');
```

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
```

This will output:

```csv
Jane,Doe
```

#### Adding layout

Under the hood this package uses the [openspout/openspout](https://github.com/openspout/openspout) package. That package contains a `StyleBuilder` that you can use to format rows. Styles can only be used on excel documents.

```php
use OpenSpout\Writer\Common\Creator\Style\StyleBuilder;
use OpenSpout\Common\Entity\Style\Color;

$style = (new StyleBuilder())
   ->setFontBold()
   ->setFontSize(15)
   ->setFontColor(Color::BLUE)
   ->setShouldWrapText()
   ->setBackgroundColor(Color::YELLOW)
   ->build();

$writer->addRow(['values', 'of', 'the', 'row'], $style);
```
To style your HeaderRow simply call the `setHeaderStyle($style)` Method.

```php
$writer->setHeaderStyle($style);
```

For more information on styles head over to [the Spout docs](https://github.com/openspout/openspout/tree/3.x/docs).

#### Creating an additional sheets

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

$writerWithAutomaticHeader->getNumberOfRows(); // returns 2
```

#### Disable BOM

You can also disable adding a BOM to the start of the file. BOM must be disabled on create and cannot be disabled after creation of the writer.

A BOM, or byte order mark, indicates a number of things for the file being written including the file being unicode as well as it's UTF encoding type.

```php
SimpleExcelWriter::createWithoutBom($this->pathToCsv, $type);
```

Additional information about BOM can be found [here](https://en.wikipedia.org/wiki/Byte_order_mark).

#### Manually working with the writer object

Under the hood this package uses the [openspout/openspout](https://github.com/openspout/openspout) package. You can get to the underlying writer that implements `\OpenSpout\Reader\WriterInterface` by calling the `getWriter` method.

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

Please see [CONTRIBUTING](https://github.com/spatie/.github/blob/main/CONTRIBUTING.md) for details.

### Security

If you've found a bug regarding security please mail [security@spatie.be](mailto:security@spatie.be) instead of using the issue tracker.

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
