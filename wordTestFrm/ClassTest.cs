using Aspose.Words;
using Aspose.Words.Layout;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace wordTestFrm
{
    class ClassTest
    {
        public void LayoutEnumerator(string SavePath)
        {
            // Open a document that contains a variety of layout entities
            // Layout entities are pages, cells, rows, lines and other objects included in the LayoutEntityType enum
            // They are defined visually by the rectangular space that they occupy in the document
            Document doc = new Document(SavePath);

            // Create an enumerator that can traverse these entities like a tree
            LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
            //Assert.AreEqual(doc, layoutEnumerator.Document);

            layoutEnumerator.MoveParent(LayoutEntityType.Page);
            //Assert.AreEqual(LayoutEntityType.Page, layoutEnumerator.Type);
            //Assert.Throws<InvalidOperationException>(() => Console.WriteLine(layoutEnumerator.Text));

            // We can call this method to make sure that the enumerator points to the very first entity before we go through it forwards
            layoutEnumerator.Reset();

            // "Visual order" means when moving through an entity's children that are broken across pages,
            // page layout takes precedence and we avoid elements in other pages and move to others on the same page
            Console.WriteLine("Traversing from first to last, elements between pages separated:");
            TraverseLayoutForward(layoutEnumerator, 1);

            // Our enumerator is conveniently at the end of the collection for us to go through the collection backwards
            Console.WriteLine("Traversing from last to first, elements between pages separated:");
            TraverseLayoutBackward(layoutEnumerator, 1);

            // "Logical order" means when moving through an entity's children that are broken across pages, 
            // node relationships take precedence
            Console.WriteLine("Traversing from first to last, elements between pages mixed:");
            TraverseLayoutForwardLogical(layoutEnumerator, 1);

            Console.WriteLine("Traversing from last to first, elements between pages mixed:");
            TraverseLayoutBackwardLogical(layoutEnumerator, 1);
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Visual" order.
        /// </summary>
        private static void TraverseLayoutForward(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveFirstChild())
                {
                    TraverseLayoutForward(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MoveNext());
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Visual" order.
        /// </summary>
        private static void TraverseLayoutBackward(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveLastChild())
                {
                    TraverseLayoutBackward(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MovePrevious());
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Logical" order.
        /// </summary>
        private static void TraverseLayoutForwardLogical(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveFirstChild())
                {
                    TraverseLayoutForwardLogical(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MoveNextLogical());
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Logical" order.
        /// </summary>
        private static void TraverseLayoutBackwardLogical(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveLastChild())
                {
                    TraverseLayoutBackwardLogical(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MovePreviousLogical());
        }

        /// <summary>
        /// Print information about layoutEnumerator's current entity to the console, indented by a number of tab characters specified by indent.
        /// The rectangle that we process at the end represents the area and location thereof that the element takes up in the document.
        /// </summary>
        private static void PrintCurrentEntity(LayoutEnumerator layoutEnumerator, int indent)
        {
            string tabs = new string('\t', indent);

            Console.WriteLine(layoutEnumerator.Kind == string.Empty
                ? $"{tabs}-> Entity type: {layoutEnumerator.Type}"
                : $"{tabs}-> Entity type & kind: {layoutEnumerator.Type}, {layoutEnumerator.Kind}");

            // Only spans can contain text
            if (layoutEnumerator.Type == LayoutEntityType.Span)
                Console.WriteLine($"{tabs}   Span contents: \"{layoutEnumerator.Text}\"");

            RectangleF leRect = layoutEnumerator.Rectangle;
            Console.WriteLine($"{tabs}   Rectangle dimensions {leRect.Width}x{leRect.Height}, X={leRect.X} Y={leRect.Y}");
            Console.WriteLine($"{tabs}   Page {layoutEnumerator.PageIndex}");
        }
    }
}
