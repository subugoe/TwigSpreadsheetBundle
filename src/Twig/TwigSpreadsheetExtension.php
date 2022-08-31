<?php

namespace MewesK\TwigSpreadsheetBundle\Twig;

use InvalidArgumentException;
use MewesK\TwigSpreadsheetBundle\Helper\Arrays;
use MewesK\TwigSpreadsheetBundle\Twig\NodeVisitor\MacroContextNodeVisitor;
use MewesK\TwigSpreadsheetBundle\Twig\NodeVisitor\SyntaxCheckNodeVisitor;
use MewesK\TwigSpreadsheetBundle\Twig\TokenParser\AlignmentTokenParser;
use MewesK\TwigSpreadsheetBundle\Twig\TokenParser\CellTokenParser;
use MewesK\TwigSpreadsheetBundle\Twig\TokenParser\DocumentTokenParser;
use MewesK\TwigSpreadsheetBundle\Twig\TokenParser\DrawingTokenParser;
use MewesK\TwigSpreadsheetBundle\Twig\TokenParser\HeaderFooterTokenParser;
use MewesK\TwigSpreadsheetBundle\Twig\TokenParser\RowTokenParser;
use MewesK\TwigSpreadsheetBundle\Twig\TokenParser\SheetTokenParser;
use MewesK\TwigSpreadsheetBundle\Wrapper\HeaderFooterWrapper;
use MewesK\TwigSpreadsheetBundle\Wrapper\PhpSpreadsheetWrapper;
use Twig\Error\RuntimeError;
use Twig\Extension\AbstractExtension;
use Twig\TwigFunction;

class TwigSpreadsheetExtension extends AbstractExtension
{
    public function __construct(private array $attributes = [])
    {
    }

    public function getAttributes(): array
    {
        return $this->attributes;
    }

    /**
     * {@inheritdoc}
     */
    public function getFunctions(): array
    {
        return [
            new TwigFunction('xlsmergestyles', fn (array $style1, array $style2): array => $this->mergeStyles($style1, $style2)),
            new TwigFunction('xlscellindex', fn (array $context): ?int => $this->getCurrentColumn($context), ['needs_context' => true]),
            new TwigFunction('xlsrowindex', fn (array $context): ?int => $this->getCurrentRow($context), ['needs_context' => true]),
        ];
    }

    /**
     * {@inheritdoc}
     *
     * @throws InvalidArgumentException
     */
    public function getTokenParsers(): array
    {
        return [
            new AlignmentTokenParser([], HeaderFooterWrapper::ALIGNMENT_CENTER),
            new AlignmentTokenParser([], HeaderFooterWrapper::ALIGNMENT_LEFT),
            new AlignmentTokenParser([], HeaderFooterWrapper::ALIGNMENT_RIGHT),
            new CellTokenParser(),
            new DocumentTokenParser($this->attributes),
            new DrawingTokenParser(),
            new HeaderFooterTokenParser([], HeaderFooterWrapper::BASETYPE_FOOTER),
            new HeaderFooterTokenParser([], HeaderFooterWrapper::BASETYPE_HEADER),
            new RowTokenParser(),
            new SheetTokenParser(),
        ];
    }

    /**
     * {@inheritdoc}
     */
    public function getNodeVisitors(): array
    {
        return [
            new MacroContextNodeVisitor(),
            new SyntaxCheckNodeVisitor(),
        ];
    }

    public function mergeStyles(array $style1, array $style2): array
    {
        return Arrays::mergeRecursive($style1, $style2);
    }

    /**
     * @throws RuntimeError
     */
    public function getCurrentColumn(array $context): ?int
    {
        if (!isset($context[PhpSpreadsheetWrapper::INSTANCE_KEY])) {
            throw new RuntimeError('The PhpSpreadsheetWrapper instance is missing.');
        }

        return $context[PhpSpreadsheetWrapper::INSTANCE_KEY]->getCurrentColumn();
    }

    /**
     * @throws RuntimeError
     */
    public function getCurrentRow(array $context): ?int
    {
        if (!isset($context[PhpSpreadsheetWrapper::INSTANCE_KEY])) {
            throw new RuntimeError('The PhpSpreadsheetWrapper instance is missing.');
        }

        return $context[PhpSpreadsheetWrapper::INSTANCE_KEY]->getCurrentRow();
    }
}
