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

it('can work with an file that has headers', function (string $extension) {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension))
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
})->with(['csv', 'ods', 'xlsx']);

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

it('can retrieve the headers', function (string $extension) {
    $headers = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension))
        ->getHeaders();

    expect($headers)->toEqual([
        0 => 'email',
        1 => 'first_name',
        2 => 'last_name',
    ]);
})->with(['csv', 'ods', 'xlsx']);

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

it('can ignore the headers', function (string $extension) {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension))
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
})->with(['csv', 'ods', 'xlsx']);

it("doesn't return headers when headers are ignored", function (string $extension) {
    $headers = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension))
        ->noHeaderRow()
        ->getHeaders();

    expect($headers)->toBeNull();
})->with(['csv', 'ods', 'xlsx']);

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

it('can use custom headers with header', function (string $extension) {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension))
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
})->with(['csv', 'ods', 'xlsx']);

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

it('can retrieve the custom header with headers', function (string $extension) {
    $headers = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension))
        ->useHeaders(['email_address', 'given_name', 'surname'])
        ->getHeaders();

    expect($headers)->toEqual([
        0 => 'email_address',
        1 => 'given_name',
        2 => 'surname',
    ]);
})->with(['csv', 'ods', 'xlsx']);

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

it('can retrieve the original headers with custom headers', function (string $extension) {
    $reader = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension))
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
})->with(['csv', 'ods', 'xlsx']);

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

it('can use an offset', function (string $extension) {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension))
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
})->with(['csv', 'ods', 'xlsx']);

it('can take a limit', function (string $extension) {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension))
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
})->with(['csv', 'ods', 'xlsx']);

it('can call `getRows` twice', function (string $extension) {
    $reader = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension));
    $firstRow = $reader->getRows()->first();
    $firstRowAgain = $reader->getRows()->first();

    expect([$firstRow, $firstRowAgain])->not->toBeNull();
})->with(['csv', 'ods', 'xlsx']);

it('can call `getRows` after `getHeaders`', function (string $extension) {
    $reader = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension));

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
})->with(['csv', 'ods', 'xlsx']);

it('can call `first` on the collection twice', function (string $extension) {
    $reader = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension));
    $rowCollection = $reader->getRows();
    $firstRow = $rowCollection->first();
    $firstRowAgain = $rowCollection->first();

    expect([$firstRow, $firstRowAgain])->not->toBeNull();
})->with(['csv', 'ods', 'xlsx']);

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

it('can use custom header row formatter', function (string $extension) {
    $rows = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension))
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
})->with(['csv', 'ods', 'xlsx']);

it('can retrieve headers with a custom formatter', function (string $extension) {
    $headers = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension))
        ->formatHeadersUsing(function ($header) {
            return $header . '_suffix';
        })
        ->getHeaders();

    expect($headers)->toEqual([
        0 => 'email_suffix',
        1 => 'first_name_suffix',
        2 => 'last_name_suffix',
    ]);
})->with(['csv', 'ods', 'xlsx']);

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

it('can use a custom encoding', function () {
    $rows = SimpleExcelReader::create(getStubPath('shift-jis-encoding.csv'))
        ->useEncoding('SHIFT_JIS')
        ->getRows()
        ->toArray();

    expect($rows)->toEqual([
        ['お名前' => '太郎', 'お名前（フリガナ）' => 'タロウ'],
    ]);
});

it('can preserve date formatting', function () {
    $reader = SimpleExcelReader::create(getStubPath('formatted_dates.xlsx'));

    $defaultDates = $reader->getRows()->pluck('created_at')->Toarray();

    expect($defaultDates[0])->toBeInstanceOf(DateTimeImmutable::class);
    expect($defaultDates[1])->toBeInstanceOf(DateTimeImmutable::class);

    $formattedDates = $reader
        ->preserveDateTimeFormatting()
        ->getRows()
        ->pluck('created_at')
        ->toArray();

    expect($formattedDates[0])->toEqual('9/20/2024');
    expect($formattedDates[1])->toEqual('9/19/2024');
});

it('can preserve empty rows', function () {
    $reader = SimpleExcelReader::create(getStubPath('empty_rows.xlsx'));

    expect($reader->getRows()->count())->toBe(2);
    expect($reader->preserveEmptyRows()->getRows()->count())->toBe(3);
});

it('can count and take rows in an all file types', function (string $extension) {
    $reader = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension));

    $lazyCollection = $reader->getRows();

    expect($lazyCollection->count())->toBe(2);
    expect($lazyCollection->take(1)->all())->toEqual([
        [
            'email' => 'john@example.com',
            'first_name' => 'john',
            'last_name' => 'doe',
        ],
    ]);
})->with(['csv', 'xlsx']);

it('can count and take rows is broken in ods', function (string $extension) {
    $reader = SimpleExcelReader::create(getStubPath('header-and-rows.'.$extension));

    $lazyCollection = $reader->getRows();

    expect($lazyCollection->count())->toBe(2);
    expect($lazyCollection->take(1)->all())->toEqual([]);
})->with(['ods']);
