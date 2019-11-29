<?php

namespace Spatie\SimpleExcel;

use Box\Spout\Common\Entity\Row;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Reader\ReaderInterface;
use Illuminate\Support\LazyCollection;

class SimpleExcelReader
{
    /** @var string */
    private $path;

    /** @var \Box\Spout\Reader\ReaderInterface */
    private $reader;

    /** @var \Box\Spout\Reader\IteratorInterface */
    private $rowIterator;

    private $processHeader = true;

    public static function create(string $file)
    {
        return new static($file);
    }

    public function __construct(string $path)
    {
        $this->path = $path;

        $this->reader = ReaderEntityFactory::createReaderFromFile($this->path);
    }

    public function getPath(): string
    {
        return $this->path;
    }

    public function noHeaderRow()
    {
        $this->processHeader = false;

        return $this;
    }

    public function useDelimiter(string $delimiter)
    {
        $this->reader->setFieldDelimiter($delimiter);

        return $this;
    }

    public function useFieldEnclosure(string $fieldEnclosure)
    {
        $this->reader->setFieldEnclosure($fieldEnclosure);

        return $this;
    }

    public function getReader(): ReaderInterface
    {
        return $this->reader;
    }

    public function getRows(): LazyCollection
    {
        $this->reader->open($this->path);

        $this->reader->getSheetIterator()->rewind();

        $sheet = $this->reader->getSheetIterator()->current();

        $this->rowIterator = $sheet->getRowIterator();

        $this->rowIterator->rewind();

        /** @var \Box\Spout\Common\Entity\Row $firstRow */
        $firstRow = $this->rowIterator->current();

        if (is_null($firstRow)) {
            $this->noHeaderRow();
        }

        if ($this->processHeader) {
            $this->headers = $firstRow->toArray();
            $this->rowIterator->next();
        }

        return LazyCollection::make(function () {
            while ($this->rowIterator->valid()) {
                $row = $this->rowIterator->current();

                yield $this->getValueFromRow($row);

                $this->rowIterator->next();
            }
        });
    }

    protected function getValueFromRow(Row $row): array
    {
        if (! $this->processHeader) {
            return $row->toArray();
        }

        $values = array_slice($row->toArray(), 0, count($this->headers));

        while (count($values) < count($this->headers)) {
            $values[] = '';
        }

        return array_combine($this->headers, $values);
    }

    public function close()
    {
        $this->reader->close();
    }

    public function __destruct()
    {
        $this->close();
    }
}
