<?php

namespace MewesK\TwigSpreadsheetBundle\Tests\Twig;

use MewesK\TwigSpreadsheetBundle\Helper\Filesystem;
use MewesK\TwigSpreadsheetBundle\Twig\TwigSpreadsheetExtension;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PHPUnit\Framework\TestCase;
use Symfony\Bridge\Twig\AppVariable;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\RequestStack;
use Twig\Environment;

abstract class BaseTwigTest extends TestCase
{
    /**
     * @var string
     */
    public const CACHE_PATH = './../../var/cache';

    /**
     * @var string
     */
    public const RESULT_PATH = './../../var/result';

    /**
     * @var string
     */
    public const RESOURCE_PATH = './Fixtures/views';

    /**
     * @var string
     */
    public const TEMPLATE_PATH = './Fixtures/templates';

    protected static Environment $environment;

    /**
     * {@inheritdoc}
     *
     * @throws \Twig\Error\LoaderError
     */
    public static function setUpBeforeClass(): void
    {
        $cachePath = sprintf('%s/%s/%s', __DIR__, static::CACHE_PATH, str_replace('\\', \DIRECTORY_SEPARATOR, static::class));

        // remove temp files
        Filesystem::remove($cachePath);
        Filesystem::remove(sprintf('%s/%s/%s', __DIR__, static::RESULT_PATH, str_replace('\\', \DIRECTORY_SEPARATOR, static::class)));

        // set up Twig environment
        $twigFileSystem = new \Twig\Loader\FilesystemLoader([sprintf('%s/%s', __DIR__, static::RESOURCE_PATH)]);
        $twigFileSystem->addPath(sprintf('%s/%s', __DIR__, static::TEMPLATE_PATH), 'templates');

        static::$environment = new Environment($twigFileSystem, ['debug' => true, 'strict_variables' => true]);
        static::$environment->addExtension(new TwigSpreadsheetExtension([
            'pre_calculate_formulas' => true,
            'cache' => [
                'bitmap' => $cachePath.'/spreadsheet/bitmap',
                'xml' => false,
            ],
            'csv_writer' => [
                'delimiter' => ',',
                'enclosure' => '"',
                'excel_compatibility' => false,
                'include_separator_line' => false,
                'line_ending' => \PHP_EOL,
                'sheet_index' => 0,
                'use_bom' => true,
            ],
        ]));
        static::$environment->setCache($cachePath.'/twig');
    }

    /**
     * @throws \Twig\Error\SyntaxError
     * @throws \Twig\Error\LoaderError
     * @throws \Symfony\Component\Filesystem\Exception\IOException
     * @throws \Twig\Error\RuntimeError
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    protected function getDocument(string $templateName, string $format): Spreadsheet|string
    {
        $format = strtolower($format);

        // prepare global variables
        $request = new Request();
        $request->setRequestFormat($format);

        $requestStack = new RequestStack();
        $requestStack->push($request);

        $appVariable = new AppVariable();
        $appVariable->setRequestStack($requestStack);

        // generate source from template
        $source = static::$environment->load($templateName.'.twig')->render(['app' => $appVariable]);

        // create path
        $resultPath = sprintf('%s/%s/%s/%s.%s', __DIR__, static::RESULT_PATH, str_replace('\\', \DIRECTORY_SEPARATOR, static::class), $templateName, $format);

        // save source
        Filesystem::dumpFile($resultPath, $source);

        // load source or return path for PDFs
        return 'pdf' === $format
            ? $resultPath
            : IOFactory::createReader(ucfirst($format))
                ->load($resultPath);
    }
}
