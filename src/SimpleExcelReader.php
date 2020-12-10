<?php

namespace Spatie\SimpleExcel;

use Box\Spout\Common\Entity\Row;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Reader\Common\Creator\ReaderFactory;
use Box\Spout\Reader\IteratorInterface;
use Box\Spout\Reader\ReaderInterface;
use Illuminate\Support\LazyCollection;

class SimpleExcelReader
{
    private string $path;

    private ReaderInterface $reader;

    private IteratorInterface $rowIterator;

    private bool $processHeader = true;

    private int $skip = 0;

    private int $limit = 0;

    private bool $useLimit = false;

    public static function create(string $file, ?string $type = null)
    {
        return new static($file, $type);
    }

    public function __construct(string $path, ?string $type = null)
    {
        $this->path = $path;

        if ($type) {
            $this->reader = ReaderFactory::createFromType($type);
        } else {
            $this->reader = ReaderEntityFactory::createReaderFromFile($this->path);
        }
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

    public function skip(int $count): SimpleExcelReader
    {
        $this->skip = $count;

        return $this;
    }

    public function take(int $count): SimpleExcelReader
    {
        $this->limit = $count;
        $this->useLimit = true;

        return $this;
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
            while ($this->rowIterator->valid() && $this->skip && $this->skip--) {
                $this->rowIterator->next();
            }
            while ($this->rowIterator->valid() && (! $this->useLimit || $this->limit--)) {
                $row = $this->rowIterator->current();

                yield $this->getValueFromRow($row);

                $this->rowIterator->next();
            }
        });
    }

    protected function getValueFromRow(Row $row): array
    {
        $values = $row->toArray();
        ksort($values);

        if (! $this->processHeader) {
            return $values;
        }

        $values = array_slice($values, 0, count($this->headers));

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
