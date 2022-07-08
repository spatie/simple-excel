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

    private static ?SimpleExcelWriter $instance = \null;

    private WriterInterface $writer;

    private string $path = '';

    private bool $processHeader = true;

    private bool $processingFirstRow = true;

    private int $numberOfRows = 0;

    public static function getInstance()
    {
        if (! is_object(self::$instance)) {
            self::$instance = new static();
        }

        return self::$instance;
    }

    private static function clearInstance(): void
    {
        self::$instance = \null;
    }

    public static function create(string $file, ?string $type = \null): self
    {
        $extension = strtolower(pathinfo($file, PATHINFO_EXTENSION));

        if ($type !== \null) {
            $extension = strtolower($type);
        }

        return match ($extension) {
            'csv' => self::getInstance()->openCsv($file),
            'ods' => self::getInstance()->openOds($file),
            'xlsx' => self::getInstance()->openXlsx($file),
            default => throw new UnsupportedTypeException('No readers supporting the given type: '.$extension),
        };
    }

    private function openCsv(string $file): self
    {
        $this->path = $file;

        $this->writer = new CsvWriter($this->getCsvOptions());

        $this->writer->openToFile($file);

        self::clearInstance();

        return $this;
    }

    private function openOds(string $file): self
    {
        $this->path = $file;

        $this->writer = new OdsWriter();

        $this->writer->openToFile($file);

        self::clearInstance();

        return $this;
    }

    private function openXlsx(string $file): self
    {
        $this->path = $file;

        $this->writer = new XlsxWriter($this->getXlsxOptions());

        $this->writer->openToFile($file);

        self::clearInstance();

        return $this;
    }

    public static function streamDownload(string $downloadName): self
    {
        $simpleExcelWriter = self::create($downloadName);

        $writer = $simpleExcelWriter->getWriter();

        $writer->openToBrowser($downloadName);

        self::clearInstance();

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

        self::clearInstance();
    }
}
