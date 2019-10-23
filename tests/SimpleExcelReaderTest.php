<?php

namespace Spatie\SimpleExcel\Tests;

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
    public function the_reader_is_macroable()
    {
        SimpleExcelReader::macro('onlyJohns', function() {
            return $this
                ->getRows()
                ->filter(function(array $row) {
                    return $row['first_name'] === 'john';
                })
                ->toArray();
        });

        $rows = SimpleExcelReader::create($this->getStubPath('header-and-rows.csv'))->onlyJohns();

        $this->assertEquals([
            [
                'email' => 'john@example.com',
                'first_name' => 'john',
                'last_name' => 'doe',
            ],
        ], $rows);
    }
}
