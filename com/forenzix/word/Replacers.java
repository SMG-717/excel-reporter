package com.forenzix.word;

import java.util.AbstractMap.SimpleEntry;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFSDT;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import com.mendix.systemwideinterfaces.core.IContext;

/**
 * This class contains functions for handling {@code Replacer} objects and 
 * applying their effects on the given Apache POI objects.
 * 
 * @author SMG
 * @see Replacer
 */
public final class Replacers {
    
    /**
     * Traverse through the body elements of the specified document and apply every
     * {@code Replacer} object from the specified list in order. If not all replacers
     * have been applied onto the document, an error is raised.
     * <p>
     * This function relies on the output of {@link Extractor#generateReplacers
     * (XWPFDocument, Claim, IContext) Extractor.generateReplacers}.
     * 
     * @param doc the target document with replaceable tags
     * @param replacers the list of replacers that need to be applied
     */
    public static void orderedReplace(XWPFDocument doc, List<Replacer> replacers) {
        final int count;
        if ((count = orderedReplace(doc.getBodyElements(), replacers)) < replacers.size()) {
            throw new RuntimeException("The replacers list has not been fully exhausted: " + 
                (replacers.size() - count) + "/" + replacers.size() + " remain.");
        }
    }

    /**
     * Traverse through the specified body elements apply as many {@code Replacer} 
     * objects from the specified list in order. The body elements can be children
     * of any IBody element, which allows this function to be called recursively
     * from within complex nodes in the document tree. A count is kept of all the
     * replacers consumed so that a potential recursive callee can make proper
     * adjustments or checks.
     * <p>
     * This function relies on the output of {@link Extractor#generateReplacers
     * (XWPFDocument, Claim, IContext) Extractor.generateReplacers}.
     * 
     * @param elements the list of body elements to traverse with replaceable tags
     * @param replacers the list of replacers that need to be applied
     */
    private static int orderedReplace(List<IBodyElement> elements, List<Replacer> replacers) {
        /*
         * The algorithm for traversing the elements and replacers simultaneously
         * is as follows:
         * 
         * r is the current Replacer
         * e is the current IBodyElement
         * While we still have elements and replacers:
         *     If e is a simple element and r can be found in it:
         *         replace r in e, then consume r
         *     Else if e is a complex element:
         *         ce is the list of terminal nodes in e  (e.g. cells in a table)
         *         For every element c in ce:
         *             rs is the list of remaining replacers starting from r
         *             recursively call orderedReplace on c and rs
         *             n is the number of replacers consumed in the call
         *             consume r, n times
         *         consume e
         *     Otherwise:  
         *         consume e
         * Return index of r
         * 
         * As of now, the simple body element types are Paragraphs and Structured 
         * Document Tags (SDTs). Although SDTs are technically complex types, they
         * cannot be edited from POI. A replacer is simply consumed if found in it.
         * 
         * The only complex body elements are Tables. Every row, then cell is 
         * processed and ran through orderedReplace.
         */

        // Indexes of the current Replacer and Body Element respectively.
        int r = 0, e = 0;
        while (e < elements.size() && r < replacers.size()) {
            final IBodyElement i = elements.get(e);

            // Since SDT elements are unchangeable, we simply consume the tag if present.
            if (i instanceof XWPFSDT && ((XWPFSDT) i).getContent().getText().contains(replacers.get(r).bookmark))
                r += 1;
            
            // Otherwise, attempt to apply the replacer in the paragraph
            else if (i instanceof XWPFParagraph && paragraphReplace((XWPFParagraph) i, replacers.get(r))) 
                r += 1;
            
            // Otherwise, attempt to explore all contents of cells of the table
            else if (i instanceof XWPFTable) {
                // Keep track of the consumed replacers from recursive calls
                final SimpleEntry<String, Integer> newr = new SimpleEntry<>("r", r);

                // Notice that only a part of the replacer list is passed.
                ((XWPFTable) i).getRows().forEach(row -> row.getTableCells().forEach(
                    cell -> newr.setValue(newr.getValue() + 
                        orderedReplace(cell.getBodyElements(), replacers.subList(newr.getValue(), replacers.size())))
                ));
                    
                // Update to account for the consumed replacers
                r = newr.getValue();
                e += 1;
            }

            // Otherwise, consume the element
            else e += 1;
        }

        return r;
    }

    /**
     * Replace a single tag in the specified paragraph. The replacement is applied
     * only on the first occurance of the tag in the pararaph. This function will
     * always introduce new runs in the paragraph to represent the replacement. 
     * All new runs are copies of the first run that contained one or more
     * character from the tag. The function {@link Replacers#genRuns(XWPFRun, 
     * String) Replacers.genRuns} is used to produce the new, identically formatted
     * runs.
     * <p>
     * <i>Tags</i> are bookmark strings that represent a target for a replacer to plant
     * its replacement string in a paragraph. They can span across one or more runs
     * and/or lines. If the entire tag exists in a single run it could be anywhere
     * within the text of a run.
     * <p>
     * DISCLAIMER: This function is designed to work with Replacers that have been
     * generated by {@link Extractor#generateReplacers(XWPFDocument, Claim, IContext)
     * Extractor.generateReplacers}, otherwise, it might produce unexpected results.
     * 
     * @param p the specified paragraph where a tag could be found
     * @param replacer the replacer that will be applied on the found tag
     * @return true if the tag was found and replaced, false otherwise
     */
    private static boolean paragraphReplace(XWPFParagraph p, Replacer replacer) {

        /*
         * Important Disclaimer:
         * This function does not work well in a vacuum. It is intended to work on
         * replacers that have been generated from Extractor.generateReplacers.
         * 
         * Take for example the input paragraph with runs as follows:
         * - "Hello "
         * - "<<Wor"
         * - "<<World!>>."
         *  
         * Applying the replacer with bookmark "<<World!>>" will produce the covered
         * runs ["<<Wor", "<<World!>>"], and at the replacement stage the resulting
         * runs will be:
         * - "Hello "
         * - "Replacement!"
         * 
         * While this example is niche and clearly a typo, it is an unintended 
         * consequence. If used in tandem with the Extractor, the tag "<<World!>>"
         * would not be generated in this context, and the problem would never arise.
         * Perhaps, for a more complete Replacer, this function could be modified 
         * slightly.
         */

        final List<Integer> cover = new LinkedList<>();
        final char[] atag = replacer.bookmark.toCharArray();

         /* 
         * The algorithm for replacing the tag in a paragraph is as follows:
         * 
         * Find as big a prefix of the tag t in p as possible. This is because if 
         * the biggest part of the tag we find is not the whole of the tag, it 
         * means the tag does not exist in p, and we abort. (This will be covered
         * in more detail)
         * 
         * Find cover, start, s, and i where:
         * - cover is the list of runs which collectively contain prefix 
         * - start is the location of prefix in the first run of cover
         * - s is the length prefix
         * - i is the location of the end of prefix in cover
         * 
         * If s < len(tag):
         *     return false
         * 
         * f is the first run in cover
         * l is the last run in cover
         * text is the text content of f
         * 
         * Delete from f all characters after start
         * If len(cover) > 1:
         *     Delete from l all characters after i
         *     Delete all runs between f and l
         * 
         * Else If i != len(text)
         *     replacement += all characters in f after i
         * 
         * Generate runs after f that represent the replacement
         */
        int i = 0, s = 0, start = -1;
        for (int j = 0; j < p.getRuns().size() && s < atag.length; j += 1) {
            final char[] arun = p.getRuns().get(j).text().toCharArray();

            // Find how many characters in the run match the tag
            // This loop runs until the characters in either run or tag are depleted
            for (i = 0; i < arun.length && s < atag.length; s = atag[s] == arun[i++] ? s + 1 : 0);

            // If we have matched one or more characters, we mark this position as start
            if (s != 0) {
                cover.add(j);
                start = start == -1 ? i - s : start;
            }
            
            // Otherwise, we reset our search
            else {
                cover.clear();
                start = -1;
            }
        }

        // If the length of prefix does not match the length of the tag, abort.
        if (cover.isEmpty() || s != atag.length) 
            return false;

        final XWPFRun first = p.getRuns().get(cover.get(0));
        final String text = first.text();

        // Remove start of tag from the first run
        first.setText(text.substring(0, start), 0);
        String rep = replacer.replacement;
        
        // If more than one run 
        if (cover.size() > 1) {
            // Remove the end of the tag from the last run
            final XWPFRun last = p.getRuns().get(cover.get(cover.size() - 1));
            last.setText(last.text().substring(i), 0);
            
            // Delete all runs in between
            for (int c = cover.size() - 1; (c -= 1) > 0; p.removeRun(cover.get(c)));
        }
        
        // Otherwise, if i is not at the end of the run 
        else if (i != text.length()) { 
            // Save the characters after i from deletion
            rep += text.substring(i);
        }
        
        // Clone the first run and fill with our replacement text
        genRuns(first, rep);
        return true;
    }

    /**
     * Generate multiple runs with formatting identical to the original run specified.
     * The number of new runs generated is based on the number of lines in <i>text</i>.
     * If text contains N lines, 2N - 1 runs will be generated, one for each line
     * content, and one for each break {@code &lt;br&gt;} in between. If the provided 
     * text is blank or null, no runs will be produced.
     * <p>
     * This function uses {@link Preprocessor#cloneRun(XWPFRun, XWPFRun) Preprocessor.cloneRun}.
     *  
     * @param origin run from which new runs will be cloned
     * @param text contents of the newly generated runs
     * @throws IllegalArgumentException if the parent of origin isn't a paragraph 
     * @see {@link Preprocessor#cloneRun(XWPFRun, XWPFRun) Preprocessor.cloneRun}
     */
    private static void genRuns(XWPFRun origin, String text) {

        // The run has to be a child of a paragraph
        if (!(origin.getParent() instanceof XWPFParagraph)) {
            throw new IllegalArgumentException(
                String.format("The given origin run is a child of %s which is not a subclass of XWPFParagraph.", 
                origin.getParent().getClass().getName())
            );
        }

        // Text has to be non-null and non-blank
        if (text == null || text.isBlank()) return;

        final XWPFParagraph p = (XWPFParagraph) origin.getParent();
        final Iterator<String> itr = text.lines().iterator();
        XWPFRun run;

        // If the string is blank, no runs will be generated.
        if (!itr.hasNext()) return;

        // Generate a run for the first line of content
        run = Preprocessor.cloneRun(p.insertNewRun(p.getRuns().indexOf(origin) + 1), origin);
        run.setText(itr.next());
        while (itr.hasNext()) {
            
            // Generate a run for line break
            run = Preprocessor.cloneRun(p.insertNewRun(p.getRuns().indexOf(run) + 1), origin);
            run.setText("", 0);
            run.addBreak();
            
            // Generate a run for Nth line
            run = Preprocessor.cloneRun(p.insertNewRun(p.getRuns().indexOf(run) + 1), origin);
            run.setText(itr.next(), 0);
        }
    }
}
