<?php

use OpenSpout\Reader\CSV\Reader;
use Spatie\SimpleExcel\SimpleExcelReader;

it('can work with an empty field', function () {
    $actualCount = SimpleExcelReader::create(getStubPath('empty.csv'))
        ->getRows()
        ->count();

    expect($actualCount)->toEqual(0);
});

it('can `getHeaders` with an empty file', function () {
    $headers = SimpleExcelReader::create(getStubPath('empty.csv'))
        ->getHeaders();

    expect($headers)->toBeNull();
});

it('can work with an file that has headers', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.csv'))
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        [
            'email' => 'john@example.com',
            'first_name' => 'john',
            'last_name' => 'doe',
        ],
        [
            'email' => 'mary-jane@example.com',
            'first_name' => 'mary jane',
            'last_name' => 'doe',
        ],
    ]);
});

it('can work with a file that has only headers', function () {
    $actualCount = SimpleExcelReader::create(getStubPath('only-header.csv'))
        ->getRows()
        ->count();

    expect($actualCount)->toEqual(0);
});

it('can work with a file where the header is too short', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-too-short.csv'))
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        [
            'email' => 'john@example.com',
            'first_name' => 'john',
        ],
    ]);
});

it('can work with a file where the row is too short', function () {
    $rows = SimpleExcelReader::create(getStubPath('row-too-short.csv'))
        ->getRows()
        ->toArray();

    expect([
        [
            'email' => 'john@example.com',
            'first_name' => '',
        ],
    ])->toEqual($rows);
});

it('can retrieve the headers', function () {
    $headers = SimpleExcelReader::create(getStubPath('header-and-rows.csv'))
        ->getHeaders();

    expect($headers)->toEqual([
        0 => 'email',
        1 => 'first_name',
        2 => 'last_name',
    ]);
});

it('can read headers when header is not on the first row', function () {
    $headers = SimpleExcelReader::create(getStubPath('header-not-on-first-row.xlsx'))
        ->headerOnRow(2)
        ->getHeaders();

    expect($headers)->toMatchArray([
        0 => 'firstname',
        1 => 'lastname',
    ]);
});

it('can read content when header is not on the first row', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-not-on-first-row.xlsx'))
        ->headerOnRow(2)
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        [
            'firstname' => 'Taylor',
            'lastname' => 'Otwell',
        ],
        [
            'firstname' => 'Adam',
            'lastname' => 'Wathan',
        ],
    ]);
});

it('can ignore the headers', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.csv'))
        ->noHeaderRow()
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        [
            0 => 'email',
            1 => 'first_name',
            2 => 'last_name',
        ],
        [
            0 => 'john@example.com',
            1 => 'john',
            2 => 'doe',
        ],
        [
            0 => 'mary-jane@example.com',
            1 => 'mary jane',
            2 => 'doe',
        ],
    ]);
});

it("doesn't return headers when headers are ignored", function () {
    $headers = SimpleExcelReader::create(getStubPath('header-and-rows.csv'))
        ->noHeaderRow()
        ->getHeaders();

    expect($headers)->toBeNull();
});

it('can use custom headers without header', function () {
    $rows = SimpleExcelReader::create(getStubPath('rows-without-header.csv'))
        ->noHeaderRow()
        ->useHeaders(['email', 'first_name', 'last_name'])
        ->getRows()
        ->toArray();

    expect($rows)->toMatchArray([
        [
            'email' => 'john@example.com',
            'first_name' => 'john',
            'last_name' => 'doe',
        ],
        [
            'email' => 'mary-jane@example.com',
            'first_name' => 'mary jane',
            'last_name' => 'doe',
        ],
    ]);
});

it('can use custom headers with header', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.csv'))
        ->useHeaders(['email_address', 'given_name', 'surname'])
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        [
            'email_address' => 'john@example.com',
            'given_name' => 'john',
            'surname' => 'doe',
        ],
        [
            'email_address' => 'mary-jane@example.com',
            'given_name' => 'mary jane',
            'surname' => 'doe',
        ],
    ]);
});

it('can use custom headers with header on row', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-not-on-first-row.xlsx'))
        ->headerOnRow(2)
        ->useHeaders(['first_name', 'last_name'])
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        [
            'first_name' => 'Taylor',
            'last_name' => 'Otwell',
        ],
        [
            'first_name' => 'Adam',
            'last_name' => 'Wathan',
        ],
    ]);
});

it('can retrieve the custom header with headers', function () {
    $headers = SimpleExcelReader::create(getStubPath('header-and-rows.csv'))
        ->useHeaders(['email_address', 'given_name', 'surname'])
        ->getHeaders();

    expect($headers)->toEqual([
        0 => 'email_address',
        1 => 'given_name',
        2 => 'surname',
    ]);
});

it('can retrieve the custom headers without headers', function () {
    $headers = SimpleExcelReader::create(getStubPath('rows-without-header.csv'))
        ->noHeaderRow()
        ->useHeaders(['email_address', 'given_name', 'surname'])
        ->getHeaders();

    expect($headers)->toEqual([
        0 => 'email_address',
        1 => 'given_name',
        2 => 'surname',
    ]);
});

it('can retrieve the original headers with custom headers', function () {
    $reader = SimpleExcelReader::create(getStubPath('header-and-rows.csv'))
        ->useHeaders(['email_address', 'given_name', 'surname']);

    $headers = $reader->getHeaders();
    $originalHeaders = $reader->getOriginalHeaders();

    expect($headers)->toEqual([
        0 => 'email_address',
        1 => 'given_name',
        2 => 'surname',
    ]);

    expect($originalHeaders)->toEqual([
        0 => 'email',
        1 => 'first_name',
        2 => 'last_name',
    ]);
});

it('can user an alternative delimiter', function () {
    $rows = SimpleExcelReader::create(getStubPath('alternative-delimiter.csv'))
        ->useDelimiter(';')
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        [
            'email' => 'john@example.com',
            'first_name' => 'john',
        ],
    ]);
});

test('the reader can get the path', function () {
    $path = getStubPath('alternative-delimiter.csv');

    $reader = SimpleExcelReader::create(getStubPath('alternative-delimiter.csv'));

    expect($reader->getPath())->toEqual($path);
});

it(
    'combines headers with correct values even though they are returned in the wrong order',
    function () {
        $rows = SimpleExcelReader::create(getStubPath('columns-returned-in-wrong-order.xlsx'))
            ->getRows()
            ->toArray();

        expect($rows)->toEqual([
            [
                'id' => 11223344,
                'place' => '',
                'status' => 'yes',
            ],
            [
                'id' => 11112222,
                'place' => '',
                'status' => 'no',
            ],
        ]);
    }
);

it('can use an offset', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.csv'))
        ->skip(1)
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        [
            'email' => 'mary-jane@example.com',
            'first_name' => 'mary jane',
            'last_name' => 'doe',
        ],
    ]);
});

it('can take a limit', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.csv'))
        ->take(1)
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        [
            'email' => 'john@example.com',
            'first_name' => 'john',
            'last_name' => 'doe',
        ],
    ]);
});

it('can call `getRows` twice', function () {
    $reader = SimpleExcelReader::create(getStubPath('header-and-rows.csv'));
    $firstRow = $reader->getRows()->first();
    $firstRowAgain = $reader->getRows()->first();

    expect([$firstRow, $firstRowAgain])->not->toBeNull();
});

it('can call `getRows` after `getHeaders`', function () {
    $reader = SimpleExcelReader::create(getStubPath('header-and-rows.csv'));

    $headers = $reader->getHeaders();

    expect($headers)->toEqual([
        0 => 'email',
        1 => 'first_name',
        2 => 'last_name',
    ]);

    $rows = $reader->getRows()->toArray();

    expect($rows)->toEqual([
        [
            'email' => 'john@example.com',
            'first_name' => 'john',
            'last_name' => 'doe',
        ],
        [
            'email' => 'mary-jane@example.com',
            'first_name' => 'mary jane',
            'last_name' => 'doe',
        ],
    ]);
});

it('can call `first` on the collection twice', function () {
    $reader = SimpleExcelReader::create(getStubPath('header-and-rows.csv'));
    $rowCollection = $reader->getRows();
    $firstRow = $rowCollection->first();
    $firstRowAgain = $rowCollection->first();

    expect([$firstRow, $firstRowAgain])->not->toBeNull();
});

it('allows setting the reader type manually', function () {
    $reader = SimpleExcelReader::create('php://input', 'csv');

    expect($reader->getReader())->toBeInstanceOf(Reader::class);
});

it('can trim the header row names', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-with-spaces.csv'))
        ->trimHeaderRow()
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        [
            'email' => 'john@example.com',
            'first_name' => 'john',
            'last_name' => 'doe',
        ],
        [
            'email' => 'mary-jane@example.com ',
            'first_name' => 'mary jane',
            'last_name' => 'doe',
        ],
    ]);
});

it('can retrieve trimmed header row names', function () {
    $headers = SimpleExcelReader::create(getStubPath('header-with-spaces.csv'))
        ->trimHeaderRow()
        ->getHeaders();

    expect($headers)->toMatchArray([
        0 => 'email',
        1 => 'first_name',
        2 => 'last_name',
    ]);
});

it('can trim the header row names with alternate characters', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-with-spaces.csv'))
        ->trimHeaderRow('e')
        ->getRows()
        ->toArray();

    expect($rows)->toMatchArray([
        [
            'mail ' => 'john@example.com',
            ' first_name ' => 'john',
            ' last_nam' => 'doe',
        ],
        [
            'mail ' => 'mary-jane@example.com ',
            ' first_name ' => 'mary jane',
            ' last_nam' => 'doe',
        ],
    ]);
});

it('can convert headers to snake case', function () {
    $rows = SimpleExcelReader::create(getStubPath('headers-not-snake-case.csv'))
        ->headersToSnakeCase()
        ->getRows()
        ->toArray();

    expect($rows)->toMatchArray([
        [
            'email' => 'john@example.com',
            'first_name' => 'john',
            'last_name' => 'doe',
            'job_title' => 'male nutter',
        ],
        [
            'email' => 'mary-jane@example.com',
            'first_name' => 'mary jane',
            'last_name' => 'doe',
            'job_title' => 'female nutter',
        ],
    ]);
});

it('can retrieve headers converted to snake case', function () {
    $headers = SimpleExcelReader::create(getStubPath('headers-not-snake-case.csv'))
        ->headersToSnakeCase()
        ->getHeaders();

    expect($headers)->toMatchArray([
        0 => 'email',
        1 => 'first_name',
        2 => 'last_name',
        3 => 'job_title',
    ]);
});

it('can use custom header row formatter', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.csv'))
        ->formatHeadersUsing(function ($header) {
            return $header . '_suffix';
        })
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        [
            'email_suffix' => 'john@example.com',
            'first_name_suffix' => 'john',
            'last_name_suffix' => 'doe',
        ],
        [
            'email_suffix' => 'mary-jane@example.com',
            'first_name_suffix' => 'mary jane',
            'last_name_suffix' => 'doe',
        ],
    ]);
});

it('can retrieve headers with a custom formatter', function () {
    $headers = SimpleExcelReader::create(getStubPath('header-and-rows.csv'))
        ->formatHeadersUsing(function ($header) {
            return $header . '_suffix';
        })
        ->getHeaders();

    expect($headers)->toEqual([
        0 => 'email_suffix',
        1 => 'first_name_suffix',
        2 => 'last_name_suffix',
    ]);
});

it('can retrieve rows with a different delimiter', function () {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows-other-delimiter.csv'))
        ->useDelimiter(';')
        ->getRows()
        ->toArray();

    expect($rows)->toMatchArray([
        [
            'email' => 'john@example.com',
            'first_name' => 'john',
            'last_name' => 'doe',
        ],
        [
            'email' => 'mary-jane@example.com',
            'first_name' => 'mary jane',
            'last_name' => 'doe',
        ],
    ]);
});

it('retrieves headers with a different delimiter', function () {
    $headers = SimpleExcelReader::create(getStubPath('header-and-rows-other-delimiter.csv'))
        ->useDelimiter(';')
        ->getHeaders();

    expect($headers)->toMatchArray([
        0 => 'email',
        1 => 'first_name',
        2 => 'last_name',
    ]);
});

it('can trim empty header title', function () {
    $headers = SimpleExcelReader::create(getStubPath('empty-header-title.csv'))
        ->trimHeaderRow()
        ->getHeaders();

    expect($headers)->toMatchArray([
        0 => 'email',
        1 => '',
        2 => 'last',
    ]);
});

it('can trim empty header title with custom trim characters', function () {
    $headers = SimpleExcelReader::create(getStubPath('empty-header-title.csv'))
        ->trimHeaderRow('il ')
        ->getHeaders();

    expect($headers)->toMatchArray([
        0 => 'ema',
        1 => '',
        2 => 'ast',
    ]);
});

it('can select the sheet of an excel file', function () {
    $reader = SimpleExcelReader::create(getStubPath('multiple_sheets.xlsx'));

    expect($reader->getHeaders())->toMatchArray([
        0 => 'firstname',
        1 => 'lastname',
    ]);

    expect($reader->fromSheet(2)->getHeaders())->toMatchArray([
        0 => 'contact',
        1 => 'email',
    ]);
});

it('will not open non-existing sheets', function () {
    SimpleExcelReader::create(getStubPath('multiple_sheets.xlsx'))
        ->fromSheet(3)
        ->getHeaders();
})->throws(InvalidArgumentException::class);

it('can select the sheet of an excel file by name', function () {
    $reader = SimpleExcelReader::create(getStubPath('multiple_sheets.xlsx'));

    expect(
        $reader->fromSheetName("sheet1")->getHeaders()
    )->toEqual([
        0 => 'firstname',
        1 => 'lastname',
    ]);

    expect(
        $reader->fromSheetName("sheet2")->getHeaders()
    )->toEqual([
        0 => 'contact',
        1 => 'email',
    ]);
});

it('Can check if a sheet exists by name', function () {
    $reader = SimpleExcelReader::create(getStubPath('multiple_sheets.xlsx'));

    expect($reader->hasSheet("sheet1"))->toBeTrue();
});

it('Can check if a sheet doesn\'t exists by name', function () {
    $reader = SimpleExcelReader::create(getStubPath('multiple_sheets.xlsx'));

    expect($reader->hasSheet("sheet0"))->toBeFalse();
});

it('will not open non-existing sheets by name', function () {
    SimpleExcelReader::create(getStubPath('multiple_sheets.xlsx'))
        ->fromSheetName("sheetNotExists")
        ->getHeaders();
})->throws(InvalidArgumentException::class);
