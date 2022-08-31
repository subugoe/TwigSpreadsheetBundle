<?php

namespace MewesK\TwigSpreadsheetBundle\Wrapper;

use Twig\Environment;

class PhpSpreadsheetWrapper
{
    /**
     * @var string
     */
    public const INSTANCE_KEY = '_tsb';

    private DocumentWrapper $documentWrapper;

    private SheetWrapper $sheetWrapper;

    private RowWrapper $rowWrapper;

    private CellWrapper $cellWrapper;

    private HeaderFooterWrapper $headerFooterWrapper;

    private DrawingWrapper $drawingWrapper;

    public function __construct(array $context, Environment $environment, array $attributes = [])
    {
        $this->documentWrapper = new DocumentWrapper($context, $environment, $attributes);
        $this->sheetWrapper = new SheetWrapper($context, $environment, $this->documentWrapper);
        $this->rowWrapper = new RowWrapper($context, $environment, $this->sheetWrapper);
        $this->cellWrapper = new CellWrapper($context, $environment, $this->sheetWrapper);
        $this->headerFooterWrapper = new HeaderFooterWrapper($context, $environment, $this->sheetWrapper);
        $this->drawingWrapper = new DrawingWrapper($context, $environment, $this->sheetWrapper, $this->headerFooterWrapper, $attributes);
    }

    /**
     * Copies the PhpSpreadsheetWrapper instance from 'varargs' to '_tsb'. This is necessary for all Twig code running
     * in sub-functions (e.g. block, macro, ...) since the root context is lost. To fix the sub-context a reference to
     * the PhpSpreadsheetWrapper instance is included in all function calls.
     */
    public static function fixContext(array $context): array
    {
        if (!isset($context[self::INSTANCE_KEY]) && isset($context['varargs']) && \is_array($context['varargs'])) {
            foreach ($context['varargs'] as $arg) {
                if ($arg instanceof self) {
                    $context[self::INSTANCE_KEY] = $arg;
                    break;
                }
            }
        }

        return $context;
    }

    public function getCurrentColumn(): ?int
    {
        return $this->sheetWrapper->getColumn();
    }

    public function getCurrentRow(): ?int
    {
        return $this->sheetWrapper->getRow();
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     * @throws \RuntimeException
     */
    public function startDocument(array $properties = []): void
    {
        $this->documentWrapper->start($properties);
    }

    /**
     * @throws \RuntimeException
     * @throws \LogicException
     * @throws \InvalidArgumentException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws \Symfony\Component\Filesystem\Exception\IOException
     */
    public function endDocument(): void
    {
        $this->documentWrapper->end();
    }

    /**
     * @param int|string|null $index
     *
     * @throws \LogicException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \RuntimeException
     */
    public function startSheet($index = null, array $properties = []): void
    {
        $this->sheetWrapper->start($index, $properties);
    }

    /**
     * @throws \LogicException
     * @throws \Exception
     */
    public function endSheet(): void
    {
        $this->sheetWrapper->end();
    }

    /**
     * @throws \LogicException
     */
    public function startRow(int $index = null): void
    {
        $this->rowWrapper->start($index);
    }

    /**
     * @throws \LogicException
     */
    public function endRow(): void
    {
        $this->rowWrapper->end();
    }

    /**
     * @throws \InvalidArgumentException
     * @throws \LogicException
     * @throws \RuntimeException
     */
    public function startCell(int $index = null, array $properties = []): void
    {
        $this->cellWrapper->start($index, $properties);
    }

    /**
     * @param mixed|null $value
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setCellValue($value = null): void
    {
        $this->cellWrapper->value($value);
    }

    public function endCell(): void
    {
        $this->cellWrapper->end();
    }

    /**
     * @throws \LogicException
     * @throws \RuntimeException
     * @throws \InvalidArgumentException
     */
    public function startHeaderFooter(string $baseType, string $type = null, array $properties = []): void
    {
        $this->headerFooterWrapper->start($baseType, $type, $properties);
    }

    /**
     * @throws \LogicException
     * @throws \InvalidArgumentException
     */
    public function endHeaderFooter(): void
    {
        $this->headerFooterWrapper->end();
    }

    /**
     * @throws \InvalidArgumentException
     * @throws \LogicException
     */
    public function startAlignment(string $type = null, array $properties = []): void
    {
        $this->headerFooterWrapper->startAlignment($type, $properties);
    }

    /**
     * @throws \InvalidArgumentException
     * @throws \LogicException
     */
    public function endAlignment(string $value = null): void
    {
        $this->headerFooterWrapper->endAlignment($value);
    }

    /**
     * @throws \Symfony\Component\Filesystem\Exception\IOException
     * @throws \InvalidArgumentException
     * @throws \LogicException
     * @throws \RuntimeException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function startDrawing(string $path, array $properties = []): void
    {
        $this->drawingWrapper->start($path, $properties);
    }

    public function endDrawing(): void
    {
        $this->drawingWrapper->end();
    }
}
