<?php

namespace MewesK\TwigSpreadsheetBundle\Tests\Twig;

class PdfTwigTest extends BaseTwigTest
{
    public function formatProvider(): array
    {
        return [['pdf']];
    }

    /**
     * @throws \Exception
     * @dataProvider formatProvider
     */
    public function testBasic(string $format): void
    {
        $path = $this->getDocument('cellProperties', $format);

        static::assertFileExists($path, 'File does not exist');
        static::assertGreaterThan(0, filesize($path), 'File is empty');
    }
}
