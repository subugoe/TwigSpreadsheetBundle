<?php

namespace MewesK\TwigSpreadsheetBundle\Wrapper;

use Twig\Environment;

class RowWrapper extends BaseWrapper
{
    public function __construct(array $context, Environment $environment, protected SheetWrapper $sheetWrapper)
    {
        parent::__construct($context, $environment);
    }

    /**
     * @throws \LogicException
     */
    public function start(int $index = null): void
    {
        if (null === $this->sheetWrapper->getObject()) {
            throw new \LogicException();
        }

        if (null === $index) {
            $this->sheetWrapper->increaseRow();
        } else {
            $this->sheetWrapper->setRow($index);
        }
    }

    /**
     * @throws \LogicException
     */
    public function end(): void
    {
        if (null === $this->sheetWrapper->getObject()) {
            throw new \LogicException();
        }

        $this->sheetWrapper->setColumn(null);
    }
}
