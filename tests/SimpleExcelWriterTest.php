<?php

use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Writer\CSV\Writer;
use Spatie\SimpleExcel\SimpleExcelReader;
use Spatie\SimpleExcel\SimpleExcelWriter;

use function Spatie\Snapshots\assertMatchesFileSnapshot;

use Spatie\TemporaryDirectory\TemporaryDirectory;

beforeEach(function () {
    $this->temporaryDirectory = new TemporaryDirectory(__DIR__ . '/temp');

    $this->pathToCsv = $this->temporaryDirectory->path('test.csv');
    $this->pathToXlsx = $this->temporaryDirectory->path('test.xlsx');
});

it('can write a regular csv', function () {
    SimpleExcelWriter::create($this->pathToCsv)
        ->addRow([
            'first_name' => 'John',
            'last_name' => 'Doe',
        ])
        ->addRow([
            'first_name' => 'Jane',
            'last_name' => 'Doe',
        ]);

    assertMatchesFileSnapshot($this->pathToCsv);
});

test('add multiple rows', function () {
    SimpleExcelWriter::create($this->pathToCsv)
        ->addRows(
            [
                [
                    'first_name' => 'John',
                    'last_name' => 'Doe',
                ],
                [
                    'first_name' => 'Jane',
                    'last_name' => 'Doe',
                ],
            ]
        );

    assertMatchesFileSnapshot($this->pathToCsv);
});

it('can use an alternative delimiter', function () {
    SimpleExcelWriter::create(
        file: $this->pathToCsv,
        delimiter: ';'
    )
        ->addRow([
            'first_name' => 'John',
            'last_name' => 'Doe',
        ])
        ->addRow([
            'first_name' => 'Jane',
            'last_name' => 'Doe',
        ]);

    assertMatchesFileSnapshot($this->pathToCsv);
});

it('can write the header from array', function () {
    $writerWithAutomaticHeader = SimpleExcelWriter::create($this->pathToCsv)
        ->addHeader(['first_name', 'last_name'])
        ->addRow(['John', 'Doe']);

    expect($writerWithAutomaticHeader->getNumberOfRows())->toEqual(2);

    assertMatchesFileSnapshot($this->pathToCsv);
});

it('can write a csv without a header', function () {
    SimpleExcelWriter::create($this->pathToCsv)
        ->noHeaderRow()
        ->addRow([
            'first_name' => 'John',
            'last_name' => 'Doe',
        ])
        ->addRow([
            'first_name' => 'Jane',
            'last_name' => 'Doe',
        ]);

    assertMatchesFileSnapshot($this->pathToCsv);
});

it('can get the number of rows written', function () {
    $writerWithAutomaticHeader = SimpleExcelWriter::create($this->pathToCsv)
        ->addRow([
            'first_name' => 'John',
            'last_name' => 'Doe',
        ]);

    expect($writerWithAutomaticHeader->getNumberOfRows())->toEqual(2);

    $writerWithoutAutomaticHeader = SimpleExcelWriter::create($this->pathToCsv)
        ->noHeaderRow()
        ->addRow([
            'first_name' => 'John',
            'last_name' => 'Doe',
        ]);

    expect($writerWithoutAutomaticHeader->getNumberOfRows())->toEqual(1);
});

test('the writer can get the path', function () {
    $writer = SimpleExcelWriter::create($this->pathToCsv);

    expect($writer->getPath())->toEqual($this->pathToCsv);
});

it('allows setting the writer type manually', function () {
    $writer = SimpleExcelWriter::create($this->pathToCsv, 'csv');

    expect($writer->getWriter())->toBeInstanceOf(Writer::class);
});

it('can write a csv without bom', function () {
    $writer = SimpleExcelWriter::createWithoutBom($this->pathToCsv)
        ->addRow([
            'first_name' => 'Jane',
            'last_name' => 'Doe',
        ]);

    assertMatchesFileSnapshot($this->pathToCsv);
});

it('can name a xlsx sheet', function () {
    $writer = SimpleExcelWriter::create($this->pathToXlsx)
                               ->nameCurrentSheet('TestSheet');

    expect($writer->getWriter()->getCurrentSheet()->getName())->toEqual('TestSheet');
});

it('can add a xlsx sheet', function () {
    $writer = SimpleExcelWriter::create($this->pathToXlsx)
                               ->addNewSheetAndMakeItCurrent();

    expect(count($writer->getWriter()->getSheets()))->toEqual(2);
    expect($writer->getWriter()->getCurrentSheet()->getName())->toEqual('Sheet2');
});

it('can add and name a xlsx sheet', function () {
    $writer = SimpleExcelWriter::create($this->pathToXlsx)
                               ->addNewSheetAndMakeItCurrent('TestSheet');

    expect(count($writer->getWriter()->getSheets()))->toEqual(2);
    expect($writer->getWriter()->getCurrentSheet()->getName())->toEqual('TestSheet');
});

it('preserves header and row styles in xlsx exports', function (bool $explicitHeader) {
    $headerStyle = new Style();
    $headerStyle = method_exists($headerStyle, 'withFontBold')
        ? $headerStyle->withFontBold(true)
        : $headerStyle->setFontBold();
    $rowStyle = new Style();
    $rowStyle = method_exists($rowStyle, 'withFontItalic')
        ? $rowStyle->withFontItalic(true)
        : $rowStyle->setFontItalic();

    $writer = SimpleExcelWriter::create($this->pathToXlsx)
        ->setHeaderStyle($headerStyle);

    if ($explicitHeader) {
        $writer->addHeader(['name', 'age']);
    }

    $writer->addRow(['name' => 'Jane', 'age' => 42], $rowStyle);
    $writer->close();

    $zip = new ZipArchive();
    $zip->open($this->pathToXlsx);
    $sheet = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
    $styles = simplexml_load_string($zip->getFromName('xl/styles.xml'));
    $zip->close();

    expect($sheet->sheetData->row)->toHaveCount(2);

    foreach ($sheet->sheetData->row as $row) {
        expect($row->c)->toHaveCount(2);

        foreach ($row->c as $cell) {
            $styleIndex = (int) $cell['s'];
            $fontIndex = (int) $styles->cellXfs->xf[$styleIndex]['fontId'];
            $font = $styles->fonts->font[$fontIndex];

            expect(isset($font->b))->toBe((int) $row['r'] === 1);
            expect(isset($font->i))->toBe((int) $row['r'] === 2);
        }
    }

    expect(SimpleExcelReader::create($this->pathToXlsx)->getRows()->all())
        ->toBe([['name' => 'Jane', 'age' => 42]]);
})->with([false, true]);

it('combines a custom csv delimiter with disabling the bom', function () {
    $writer = SimpleExcelWriter::create(
        file: $this->pathToCsv,
        delimiter: ';',
        shouldAddBom: false,
    )->addRow(['name' => 'Jane', 'age' => 42]);
    $writer->close();

    expect(file_get_contents($this->pathToCsv))->toBe("name;age\nJane;42\n");
});
