<?php

use OpenSpout\Writer\CSV\Writer;
use Spatie\SimpleExcel\SimpleExcelWriter;
use function Spatie\Snapshots\assertMatchesFileSnapshot;

use Spatie\TemporaryDirectory\TemporaryDirectory;

beforeEach(function () {
    $this->temporaryDirectory = new TemporaryDirectory(__DIR__ . '/temp');

    $this->pathToCsv = $this->temporaryDirectory->path('test.csv');
});

it('can write a regular CSV', function () {
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
    SimpleExcelWriter::create($this->pathToCsv)
        ->useDelimiter(';')
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

it('can write a CSV without a header', function () {
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
    $writer = SimpleExcelWriter::create('php://output', 'csv');

    expect($writer->getWriter())->toBeInstanceOf(Writer::class);
});

it('can write a CSV without bom', function () {
    $writer = SimpleExcelWriter::createWithoutBom($this->pathToCsv)
        ->addRow([
            'first_name' => 'Jane',
            'last_name' => 'Doe',
        ]);

    assertMatchesFileSnapshot($this->pathToCsv);
});
