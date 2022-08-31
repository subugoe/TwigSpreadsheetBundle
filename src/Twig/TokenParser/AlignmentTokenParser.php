<?php

namespace MewesK\TwigSpreadsheetBundle\Twig\TokenParser;

use MewesK\TwigSpreadsheetBundle\Twig\Node\AlignmentNode;
use MewesK\TwigSpreadsheetBundle\Wrapper\HeaderFooterWrapper;
use Twig\Node\Node;

class AlignmentTokenParser extends BaseTokenParser
{
    private string $alignment;

    /**
     * @param array $attributes optional attributes for the corresponding node
     *
     * @throws \InvalidArgumentException
     */
    public function __construct(array $attributes = [], string $alignment = HeaderFooterWrapper::ALIGNMENT_CENTER)
    {
        parent::__construct($attributes);

        $this->alignment = HeaderFooterWrapper::validateAlignment(strtolower($alignment));
    }

    /**
     * {@inheritdoc}
     *
     * @throws \InvalidArgumentException
     */
    public function createNode(array $nodes = [], int $lineNo = 0): Node
    {
        return new AlignmentNode($nodes, $this->getAttributes(), $lineNo, $this->getTag(), $this->alignment);
    }

    /**
     * {@inheritdoc}
     */
    public function getTag(): string
    {
        return 'xls'.$this->alignment;
    }
}
