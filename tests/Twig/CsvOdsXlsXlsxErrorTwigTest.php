<?php

namespace MewesK\TwigSpreadsheetBundle\Tests\Twig;

class CsvOdsXlsXlsxErrorTwigTest extends BaseTwigTest
{
    public function formatProvider(): array
    {
        return [['csv'], ['ods'], ['xls'], ['xlsx']];
    }

    /**
     * @throws \Exception
     * @dataProvider formatProvider
     */
    public function testDocumentError(string $format): void
    {
        $this->expectException(\Twig\Error\SyntaxError::class);
        $this->expectExceptionMessage('Node "MewesK\TwigSpreadsheetBundle\Twig\Node\DocumentNode" is not allowed inside of Node "MewesK\TwigSpreadsheetBundle\Twig\Node\SheetNode"');

        $this->getDocument('documentError', $format);
    }

    /**
     * @throws \Exception
     * @dataProvider formatProvider
     */
    public function testDocumentErrorTextAfter(string $format): void
    {
        $this->expectException(\Twig\Error\SyntaxError::class);
        $this->expectExceptionMessage('Node "'.\Twig\Node\TextNode::class.'" is not allowed after Node "MewesK\TwigSpreadsheetBundle\Twig\Node\DocumentNode"');

        $this->getDocument('documentErrorTextAfter', $format);
    }

    /**
     * @throws \Exception
     * @dataProvider formatProvider
     */
    public function testDocumentErrorTextBefore(string $format): void
    {
        $this->expectException(\Twig\Error\SyntaxError::class);
        $this->expectExceptionMessage('Node "'.\Twig\Node\TextNode::class.'" is not allowed before Node "MewesK\TwigSpreadsheetBundle\Twig\Node\DocumentNode"');

        $this->getDocument('documentErrorTextBefore', $format);
    }

    /**
     * @throws \Exception
     * @dataProvider formatProvider
     */
    public function testStartCellIndexError(string $format): void
    {
        $this->expectException(\TypeError::class);

        $this->getDocument('cellIndexError', $format);
    }

    /**
     * @throws \Exception
     * @dataProvider formatProvider
     */
    public function testStartRowIndexError(string $format): void
    {
        $this->expectException(\TypeError::class);

        $this->getDocument('rowIndexError', $format);
    }

    /**
     * @throws \Exception
     * @dataProvider formatProvider
     */
    public function testSheetError(string $format): void
    {
        $this->expectException(\Twig\Error\SyntaxError::class);
        $this->expectExceptionMessage('Node "MewesK\TwigSpreadsheetBundle\Twig\Node\RowNode" is not allowed inside of Node "MewesK\TwigSpreadsheetBundle\Twig\Node\DocumentNode"');

        $this->getDocument('sheetError', $format);
    }
}
