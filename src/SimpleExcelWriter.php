<?php
declare(strict_types=1);

namespace Spatie\SimpleExcel;

use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Exception\UnsupportedTypeException;
use OpenSpout\Writer\CSV\Writer as CsvWriter;
use OpenSpout\Writer\ODS\Writer as OdsWriter;
use OpenSpout\Writer\WriterInterface;
use OpenSpout\Writer\XLSX\Writer as XlsxWriter;

class SimpleExcelWriter
{
    use SimpleExcelWriterOptions;

    private WriterInterface $writer;

    private string $path = '';

    private bool $processHeader = true;

    private bool $processingFirstRow = true;

    private int $numberOfRows = 0;

    public static function options(): self
    {
        return new static();
    }

    public static function create(string $file): self
    {
        $extension = strtolower(pathinfo($file, PATHINFO_EXTENSION));

        return match ($extension) {
            'csv' => (new self)->openCsv($file),
            'ods' => (new self)->openOds($file),
            'xlsx' => (new self)->openXlsx($file),
            default => throw new UnsupportedTypeException('No readers supporting the given type: '.$extension),
        };
    }

    public function openCsv(string $file): self
    {
        $this->path = $file;

        $this->writer = new CsvWriter($this->getCsvOptions());

        $this->writer->openToFile($file);

        return $this;
    }

    public static function createCsv(string $file): self
    {
        return (new self)->openCsv($file);
    }

    public function openOds(string $file): self
    {
        $this->path = $file;

        $this->writer = new OdsWriter();

        $this->writer->openToFile($file);

        return $this;
    }

    public static function createOds(string $file): self
    {
        return (new self)->openOds($file);
    }

    public function openXlsx(string $file): self
    {
        $this->path = $file;

        $this->writer = new XlsxWriter($this->getXlsxOptions());

        $this->writer->openToFile($file);

        return $this;
    }

    public static function createXlsx(string $file): self
    {
        return (new self)->openXlsx($file);
    }

    public static function streamDownload(string $downloadName): self
    {
        $simpleExcelWriter = self::create($downloadName);

        $writer = $simpleExcelWriter->getWriter();

        $writer->openToBrowser($downloadName);

        return $simpleExcelWriter;
    }

    public function getPath(): string
    {
        return $this->path;
    }

    public function getWriter(): WriterInterface
    {
        return $this->writer;
    }

    public function getNumberOfRows(): int
    {
        return $this->numberOfRows;
    }

    public function noHeaderRow(): self
    {
        $this->processHeader = false;

        return $this;
    }

    /**
     *  @param \OpenSpout\Common\Entity\Style\Style $style
     */
    public function setHeaderStyle(Style $style): self
    {
        $this->headerStyle = $style;

        return $this;
    }

    /**
     * @param \OpenSpout\Common\Entity\Row|array $row
     * @param \OpenSpout\Common\Entity\Style\Style|null $style
     */
    public function addRow(Row|array $row, Style $style = null): self
    {
        if (is_array($row)) {
            $rowEntity = Row::fromValues($row, $style);
        }

        if ($this->processHeader && $this->processingFirstRow) {
            $header = Row::fromValues(array_keys($row), $style);
            $this->writeHeaderFromRow($header);
        }

        $this->writer->addRow($rowEntity);
        $this->numberOfRows++;

        $this->processingFirstRow = false;

        return $this;
    }

    public function addRows(iterable $rows): self
    {
        foreach ($rows as $row) {
            $this->addRow($row);
        }

        return $this;
    }

    protected function writeHeaderFromRow(Row $row): void
    {
        $this->writer->addRow($row);
        $this->numberOfRows++;
    }

    public function toBrowser(): void
    {
        $this->writer->close();

        exit;
    }

    public function close()
    {
        $this->writer->close();
    }

    public function __destruct()
    {
        $this->close();
    }
}
