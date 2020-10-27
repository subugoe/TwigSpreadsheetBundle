<?php

namespace MewesK\TwigSpreadsheetBundle\Twig\NodeVisitor;

use MewesK\TwigSpreadsheetBundle\Twig\Node\BaseNode;
use MewesK\TwigSpreadsheetBundle\Twig\Node\DocumentNode;
use Twig\NodeVisitor\AbstractNodeVisitor;

/**
 * Class SyntaxCheckNodeVisitor.
 */
class SyntaxCheckNodeVisitor extends AbstractNodeVisitor
{
    /**
     * @var array
     */
    protected $path = [];

    /**
     * {@inheritdoc}
     */
    public function getPriority()
    {
        return 0;
    }

    /**
     * {@inheritdoc}
     *
     * @throws \Twig\Error\SyntaxError
     */
    protected function doEnterNode(\Twig\Node\Node $node, \Twig\Environment $env)
    {
        try {
            if ($node instanceof BaseNode) {
                $this->checkAllowedParents($node);
            } else {
                $this->checkAllowedChildren($node);
            }
        } catch (\Twig\Error\SyntaxError $e) {
            // reset path since throwing an error prevents doLeaveNode to be called
            $this->path = [];
            throw $e;
        }

        $this->path[] = $node !== null ? \get_class($node) : null;

        return $node;
    }

    /**
     * {@inheritdoc}
     */
    protected function doLeaveNode(\Twig\Node\Node $node, \Twig\Environment $env)
    {
        array_pop($this->path);

        return $node;
    }

    /**
     * @param \Twig\Node\Node $node
     *
     * @throws \Twig\Error\SyntaxError
     */
    private function checkAllowedChildren(\Twig\Node\Node $node)
    {
        $hasDocumentNode = false;
        $hasTextNode = false;

        /**
         * @var \Twig\Node\Node $currentNode
         */
        foreach ($node->getIterator() as $currentNode) {
            if ($currentNode instanceof \Twig\Node\TextNode) {
                if ($hasDocumentNode) {
                    throw new \Twig\Error\SyntaxError(sprintf('Node "%s" is not allowed after Node "%s".', \Twig\Node\TextNode::class, DocumentNode::class));
                }
                $hasTextNode = true;
            } elseif ($currentNode instanceof DocumentNode) {
                if ($hasTextNode) {
                    throw new \Twig\Error\SyntaxError(sprintf('Node "%s" is not allowed before Node "%s".', \Twig\Node\TextNode::class, DocumentNode::class));
                }
                $hasDocumentNode = true;
            }
        }
    }

    /**
     * @param BaseNode $node
     *
     * @throws \Twig\Error\SyntaxError
     */
    private function checkAllowedParents(BaseNode $node)
    {
        $parentName = null;

        // find first parent from this bundle
        foreach (array_reverse($this->path) as $className) {
            if (strpos($className, 'MewesK\\TwigSpreadsheetBundle\\Twig\\Node\\') === 0) {
                $parentName = $className;
                break;
            }
        }

        // allow no parents (e.g. macros, includes)
        if ($parentName === null) {
            return;
        }

        // check if parent is allowed
        foreach ($node->getAllowedParents() as $className) {
            if ($className === $parentName) {
                return;
            }
        }

        throw new \Twig\Error\SyntaxError(sprintf('Node "%s" is not allowed inside of Node "%s".', \get_class($node), $parentName));
    }
}
