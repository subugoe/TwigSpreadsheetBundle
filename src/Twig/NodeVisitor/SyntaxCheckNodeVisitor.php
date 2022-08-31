<?php

namespace MewesK\TwigSpreadsheetBundle\Twig\NodeVisitor;

use MewesK\TwigSpreadsheetBundle\Twig\Node\BaseNode;
use MewesK\TwigSpreadsheetBundle\Twig\Node\DocumentNode;
use Twig\Environment;
use Twig\Error\SyntaxError;
use Twig\Node\Node;
use Twig\Node\TextNode;
use Twig\NodeVisitor\AbstractNodeVisitor;

class SyntaxCheckNodeVisitor extends AbstractNodeVisitor
{
    protected array $path = [];

    /**
     * {@inheritdoc}
     */
    public function getPriority(): int
    {
        return 0;
    }

    /**
     * {@inheritdoc}
     *
     * @throws SyntaxError
     */
    protected function doEnterNode(Node $node, Environment $env): Node
    {
        try {
            if ($node instanceof BaseNode) {
                $this->checkAllowedParents($node);
            } else {
                $this->checkAllowedChildren($node);
            }
        } catch (SyntaxError $syntaxError) {
            // reset path since throwing an error prevents doLeaveNode to be called
            $this->path = [];
            throw $syntaxError;
        }

        $this->path[] = null !== $node ? $node::class : null;

        return $node;
    }

    /**
     * {@inheritdoc}
     */
    protected function doLeaveNode(Node $node, Environment $env): ?Node
    {
        array_pop($this->path);

        return $node;
    }

    /**
     * @throws SyntaxError
     * @throws \Exception
     */
    private function checkAllowedChildren(Node $node): void
    {
        $hasDocumentNode = false;
        $hasTextNode = false;

        /**
         * @var Node $currentNode
         */
        foreach ($node->getIterator() as $currentNode) {
            if ($currentNode instanceof TextNode) {
                if ($hasDocumentNode) {
                    throw new SyntaxError(sprintf('Node "%s" is not allowed after Node "%s".', TextNode::class, DocumentNode::class));
                }

                $hasTextNode = true;
            } elseif ($currentNode instanceof DocumentNode) {
                if ($hasTextNode) {
                    throw new SyntaxError(sprintf('Node "%s" is not allowed before Node "%s".', TextNode::class, DocumentNode::class));
                }

                $hasDocumentNode = true;
            }
        }
    }

    /**
     * @throws SyntaxError
     */
    private function checkAllowedParents(BaseNode $node): void
    {
        $parentName = null;

        // find first parent from this bundle
        foreach (array_reverse($this->path) as $className) {
            if (str_starts_with($className, 'MewesK\\TwigSpreadsheetBundle\\Twig\\Node\\')) {
                $parentName = $className;
                break;
            }
        }

        // allow no parents (e.g. macros, includes)
        if (null === $parentName) {
            return;
        }

        // check if parent is allowed
        foreach ($node->getAllowedParents() as $className) {
            if ($className === $parentName) {
                return;
            }
        }

        throw new SyntaxError(sprintf('Node "%s" is not allowed inside of Node "%s".', $node::class, $parentName));
    }
}
