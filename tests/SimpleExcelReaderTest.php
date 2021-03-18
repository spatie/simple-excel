<?php

namespace Spatie\SimpleExcel\Tests;

use Box\Spout\Reader\CSV\Reader;
use Spatie\SimpleExcel\SimpleExcelReader;

class SimpleExcelReaderTest extends TestCase
{
    /** @test */
    public function it_can_work_with_an_empty_file()
    {
        $actualCount = SimpleExcelReader::create($this->getStubPath('empty.csv'))
            ->getRows()
            ->count();

        $this->assertEquals(0, $actualCount);
    }

    /** @test */
    public function it_can_getHeaders_with_an_empty_file()
    {
        $headers = SimpleExcelReader::create($this->getStubPath('empty.csv'))
            ->getHeaders();

        $this->assertEquals(null, $headers);
    }

    /** @test */
    public function it_can_work_with_an_file_that_has_headers()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'))
            ->getRows()
            ->toArray();

        $this->assertEquals([
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
        ], $rows);
    }

    /** @test */
    public function it_can_work_with_a_file_that_has_only_headers()
    {
        $actualCount = SimpleExcelReader::create($this->getStubPath('only-header.csv'))
            ->getRows()
            ->count();

        $this->assertEquals(0, $actualCount);
    }

    /** @test */
    public function it_can_work_with_a_file_where_the_header_is_too_short()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('header-too-short.csv'))
            ->getRows()
            ->toArray();

        $this->assertEquals([
            [
                'email' => 'john@example.com',
                'first_name' => 'john',
            ],
        ], $rows);
    }

    /** @test */
    public function it_can_work_with_a_file_where_the_row_is_too_short()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('row-too-short.csv'))
            ->getRows()
            ->toArray();

        $this->assertEquals([
            [
                'email' => 'john@example.com',
                'first_name' => '',
            ],
        ], $rows);
    }

    /** @test */
    public function it_can_retrieve_the_headers()
    {
        $headers = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'))
            ->getHeaders();

        $this->assertEquals([
            0 => 'email',
            1 => 'first_name',
            2 => 'last_name',
        ], $headers);
    }

    /** @test */
    public function it_can_ignore_the_headers()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'))
            ->noHeaderRow()
            ->getRows()
            ->toArray();

        $this->assertEquals([
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
        ], $rows);
    }

    /** @test */
    public function it_doesnt_return_headers_when_headers_are_ignored()
    {
        $headers = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'))
            ->noHeaderRow()
            ->getHeaders();

        $this->assertEquals(null, $headers);
    }

    /** @test */
    public function it_can_use_an_alternative_delimiter()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('alternative-delimiter.csv'))
            ->useDelimiter(';')
            ->getRows()
            ->toArray();

        $this->assertEquals([
            [
                'email' => 'john@example.com',
                'first_name' => 'john',
            ],
        ], $rows);
    }

    /** @test */
    public function the_reader_can_get_the_path()
    {
        $path = $this->getStubPath('alternative-delimiter.csv');

        $reader = SimpleExcelReader::create($this->getStubPath('alternative-delimiter.csv'));

        $this->assertEquals($path, $reader->getPath());
    }

    /** @test */
    public function it_combines_headers_with_correct_values_even_though_they_are_returned_in_the_wrong_order()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('columns-returned-in-wrong-order.xlsx'))
            ->getRows()
            ->toArray();

        $this->assertEquals([
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
        ], $rows);
    }

    /** @test */
    public function it_can_use_an_offset()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'))
            ->skip(1)
            ->getRows()
            ->toArray();

        $this->assertEquals([
            [
                'email' => 'mary-jane@example.com',
                'first_name' => 'mary jane',
                'last_name' => 'doe',
            ],
        ], $rows);
    }

    /** @test */
    public function it_can_take_a_limit()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'))
            ->take(1)
            ->getRows()
            ->toArray();

        $this->assertEquals([
            [
                'email' => 'john@example.com',
                'first_name' => 'john',
                'last_name' => 'doe',
            ],
        ], $rows);
    }

    /** @test */
    public function it_can_call_getRows_twice()
    {
        $reader = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'));
        $firstRow = $reader->getRows()->first();
        $firstRowAgain = $reader->getRows()->first();

        $this->assertNotNull($firstRow);
        $this->assertNotNull($firstRowAgain);
    }

    /** @test */
    public function it_can_call_getRows_after_getHeaders()
    {
        $reader = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'));

        $headers = $reader->getHeaders();

        $this->assertEquals([
            0 => 'email',
            1 => 'first_name',
            2 => 'last_name',
        ], $headers);

        $rows = $reader->getRows()->toArray();

        $this->assertEquals([
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
        ], $rows);
    }

    /** @test */
    public function it_can_call_first_on_the_collection_twice()
    {
        $reader = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'));
        $rowCollection = $reader->getRows();
        $firstRow = $rowCollection->first();
        $firstRowAgain = $rowCollection->first();

        $this->assertNotNull($firstRow);
        $this->assertNotNull($firstRowAgain);
    }

    /** @test */
    public function it_allows_setting_the_reader_type_manually()
    {
        $reader = SimpleExcelReader::create('php://input', 'csv');

        $this->assertInstanceOf(Reader::class, $reader->getReader());
    }

    /** @test */
    public function it_can_trim_the_header_row_names()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('header-with-spaces.csv'))
            ->trimHeaderRow()
            ->getRows()
            ->toArray();

        $this->assertEquals([
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
        ], $rows);
    }

    /** @test */
    public function it_can_retrieve_trimmed_header_row_names()
    {
        $headers = SimpleExcelReader::create($this->getStubPath('header-with-spaces.csv'))
            ->trimHeaderRow()
            ->getHeaders();

        $this->assertEquals([
            0 => 'email',
            1 => 'first_name',
            2 => 'last_name',
        ], $headers);
    }

    /** @test */
    public function it_can_trim_the_header_row_names_with_alternate_characters()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('header-with-spaces.csv'))
            ->trimHeaderRow('e')
            ->getRows()
            ->toArray();

        $this->assertEquals([
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
        ], $rows);
    }

    /** @test */
    public function it_can_convert_headers_to_snake_case()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('headers-not-snake-case.csv'))
            ->headersToSnakeCase()
            ->getRows()
            ->toArray();

        $this->assertEquals([
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
        ], $rows);
    }

    /** @test */
    public function it_can_retrieve_headers_converted_to_snake_case()
    {
        $headers = SimpleExcelReader::create($this->getStubPath('headers-not-snake-case.csv'))
            ->headersToSnakeCase()
            ->getHeaders();

        $this->assertEquals([
            0 => 'email',
            1 => 'first_name',
            2 => 'last_name',
            3 => 'job_title',
        ], $headers);
    }

    /** @test */
    public function it_can_use_custom_header_row_formatter()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'))
            ->formatHeadersUsing(function ($header) {
                return $header . '_suffix';
            })
            ->getRows()
            ->toArray();

        $this->assertEquals([
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
        ], $rows);
    }

    /** @test */
    public function it_can_retrieve_headers_with_a_custom_formatter()
    {
        $headers = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'))
            ->formatHeadersUsing(function ($header) {
                return $header . '_suffix';
            })
            ->getHeaders();

        $this->assertEquals([
            0 => 'email_suffix',
            1 => 'first_name_suffix',
            2 => 'last_name_suffix',
        ], $headers);
    }

    /** @test */
    public function it_can_retrieve_rows_with_a_different_delimiter()
    {
        $rows = SimpleExcelReader::create($this->getStubPath('header-and-rows-other-delimiter.csv'))
            ->useDelimiter(';')
            ->getRows()
            ->toArray();

        $this->assertEquals([
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
        ], $rows);
    }

    /** @test */
    public function it_can_retrieve_headers_with_a_different_delimiter()
    {
        $headers = SimpleExcelReader::create($this->getStubPath('header-and-rows-other-delimiter.csv'))
            ->useDelimiter(';')
            ->getHeaders();

        $this->assertEquals([
            0 => 'email',
            1 => 'first_name',
            2 => 'last_name',
        ], $headers);
    }
}
