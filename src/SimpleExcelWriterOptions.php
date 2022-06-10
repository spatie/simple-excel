<?php

declare(strict_types=1);

namespace Spatie\SimpleExcel;

use OpenSpout\Writer\CSV\Options as CsvOptions;
use OpenSpout\Writer\XLSX\Options as XlsxOptions;

trait SimpleExcelWriterOptions
{
    /* CSV options */
    public string $field_delimiter = ',';
    public string $field_enclosure = '"';
    public bool $should_add_bom = true;
    public int $flush_threshold = 500;

    /* XLSX options */
    public bool $should_use_inline_strings = true;

    public function withoutInlineStrings(): self
    {
        $this->should_use_inline_strings = \false;

        return $this;
    }

    public function flushThreshold(int $threshold): self
    {
        $this->flush_threshold = $threshold;

        return $this;
    }

    public function withoutBom(): self
    {
        $this->should_add_bom = \false;

        return $this;
    }

    public function useDelimiter(string $delimiter): self
    {
        $this->field_delimiter = $delimiter;

        return $this;
    }

    public function getCsvOptions(): CsvOptions
    {
        $options = new CsvOptions();
        $options->FIELD_DELIMITER = $this->field_delimiter;
        $options->FIELD_ENCLOSURE = $this->field_enclosure;
        $options->SHOULD_ADD_BOM = $this->should_add_bom;
        $options->FLUSH_THRESHOLD = $this->flush_threshold;

        return $options;
    }

    public function getXlsxOptions(): XlsxOptions
    {
        $options = new XlsxOptions();
        $options->SHOULD_USE_INLINE_STRINGS = $this->should_use_inline_strings;

        return $options;
    }
}
