package com.forenzix.word;

import java.util.LinkedList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFSDT;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlCursor.TokenType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;

/**
 * An Apache POI based document processor. It is made for the automatic
 * replication
 * of specific portions of an {@link XWPFDocument} object. This class has the
 * potential
 * to house many more preprocessing routines, but for now it does not do much.
 * 
 * @author SMG
 * @see XWPFDocument
 */
public final class Preprocessor {

    private final static Pattern pattern = Pattern.compile("repeat_start_(\\d+)");

    // No object instances.
    private Preprocessor() {
        throw new UnsupportedOperationException("No object instances.");
    }

    /**
     * Lock the specified document using a password. The algorithm used to hash
     * the password is Sha256.
     * 
     * @param doc      document to be locked
     * @param password plain text password
     */
    public static void lock(XWPFDocument doc, String password) {
        doc.enforceReadonlyProtection(password, HashAlgorithm.sha256);
    }
    
    /**
     * Capture all the elements between two specific tags, namely {@code "repeat_start"}
     * and {@code"repeat_end"}, and create multiple copies of them in a row.
     * The paragraphs in which these tags exist will be erased from the document
     * before the document is saved
     * 
     * @param doc   document to be processed
     * @param times the number of times for which the tagged sections shall be
     *              repeated
     */
    public static void repeatSections(XWPFDocument doc) {
        repeatSections(doc, 1);
    }

    /**
     * Capture all the elements between two specific tags, namely {@code "repeat_start"}
     * and {@code"repeat_end"}, and create multiple copies of them in a row.
     * The paragraphs in which these tags exist will be erased from the document
     * before the document is saved
     * 
     * @param doc   document to be processed
     * @param times the number of times for which the tagged sections shall be
     *              repeated
     */
    public static void repeatSections(XWPFDocument doc, long times) {
        
        int SDTCount = 0;
        final long original_times = times;
        
        // Search through the elements of the document
        for (int i = 0, start = 0; i < doc.getBodyElements().size(); i += 1) {
            final IBodyElement element = doc.getBodyElements().get(i);

            if (!(element instanceof XWPFParagraph)) {
                if (element instanceof XWPFSDT)
                    SDTCount += 1;
                continue;
            }

            // ... until a paragraph that contains the tags 'repeat start' or 'repeat end'
            // is found
            final String content = ((XWPFParagraph) element).getText().toLowerCase();
            if (content.contains("repeat_start")) {
                // Record where the repeated section (tags exclusive) starts
                start = i + 1;
                final Matcher match = pattern.matcher(content);
                if (match.find()) {
                    try {
                        times = Integer.parseInt(match.group().substring("repeat_start_".length()));
                    } catch (NumberFormatException e) {}
                }
                else {
                    times = original_times;
                }
            }

            else if (content.contains("repeat_end")) {
                
    /* 
     * NOTE: Notice in the next line that SDTCount is considered when locating 
     * where the cursor should be when copying elements over from 'start' to 
     * just before the repeat_end tag.
     * 
     * This is because I have reasons to believe that the implementation of
     * insertNewParagraph in XWPFDocument is somewhat faulty. When inserting an 
     * element at the location of the cursor in the document, it counts the 
     * number of elements before the cursor and uses that as an offset for 
     * placing the new bodyElement into the list of its siblings. The bizzare 
     * thing however is that XWPFSDT (Table of Contents) elements are not 
     * included in this tally despite being a full body element!
     * 
     * This means that for documents that include an SDT, the copied elements
     * are placed n elements higher in the element list, where n is the number 
     * of SDT elements seen before the repeat_end tag. SDTs aren't the only
     * elements that could potentially affect this, but they're simply the only
     * other objects that implement IBodyElement.
     * 
     * I adjust for this difference by selecting an element that overshoots 
     * repeat_end by a certain amount, and while this solution works for the
     * meantime, I'm not too sure about its ramifications. While the list of
     * body elements is now fixed, other lists aren't, like paragraphs and 
     * tables; they still act as if the copied element is pasted after 
     * repeat_end. And while I suspect that deleting the repeat_start and end 
     * tags after the fact rectifies things, I cannot guarantee that it would be
     * error free.
     * 
     * For now, let this wall of text be the first thing another developer
     * sees when attempting to fix this obscure issue.
     * 
     *     - SMG  
     */  
                final XmlCursor cursor = newCursor(doc.getBodyElements().get(i + SDTCount));

                for (int k = 0; k < times - 1; k += 1) {
                    // Start copying from the top down
                    for (int j = start; j < i; j += 1) {
                        final IBodyElement copyable = doc.getBodyElements().get(j);
                        
                        // And clone the copyable element accordingly
                        if (copyable instanceof XWPFParagraph) {
                            cloneParagraph(doc.insertNewParagraph(cursor), (XWPFParagraph) copyable);
                        } else if (copyable instanceof XWPFTable) {
                            cloneTable(doc.insertNewTbl(cursor), (XWPFTable) copyable);
                        }

                        while (cursor.toNextToken() != TokenType.START);
                    }
                }
                
                // Enumerate all tags
                int amount = (i - start) * ((int) times - 1);
                enumerateElements(doc.getBodyElements().subList(start, i + amount), i - start);
                
                // And remove the start and end tags
                doc.removeBodyElement(i + amount);
                doc.removeBodyElement(start - 1);

                // Advance past the copied section.
                i += amount - 2;
            }
        }
    }

    private static XmlCursor newCursor(IBodyElement e) {
        if (e instanceof XWPFParagraph) {
            return ((XWPFParagraph) e).getCTP().newCursor();
        }
        else if (e instanceof XWPFTable) {
            return ((XWPFTable) e).getCTTbl().newCursor();
        }
        else {
            return null;
        }
    }

    // Courtesy of Gary Forbis from StackOverflow
    // https://stackoverflow.com/a/23136358
    private static XWPFParagraph cloneParagraph(XWPFParagraph clone, XWPFParagraph source) {
        CTPPr pPr = clone.getCTP().isSetPPr() ? clone.getCTP().getPPr() : clone.getCTP().addNewPPr();
        pPr.set(source.getCTP().getPPr());
        for (XWPFRun r : source.getRuns()) {
            XWPFRun nr = clone.createRun();
            cloneRun(nr, r);
        }

        return clone;
    }

    // This function sometimes is used by classes other than this one. Change it
    // sparsely.
    static XWPFRun cloneRun(XWPFRun clone, XWPFRun source) {
        CTRPr rPr = clone.getCTR().isSetRPr() ? clone.getCTR().getRPr() : clone.getCTR().addNewRPr();
        rPr.set(source.getCTR().getRPr());

        final String sourceText = source.text();
        clone.setText(sourceText, 0);
        if (sourceText.isBlank()) {
            int brs = source.getCTR().getBrList().size();
            for (int i = 0; i < brs; i += 1)
                clone.getCTR().addNewBr();
        }

        return clone;
    }

    // Courtesy of... myself, SMG.
    private static XWPFTable cloneTable(XWPFTable clone, XWPFTable source) {
        CTTblPr tPr = clone.getCTTbl().getTblPr() != null ? clone.getCTTbl().getTblPr()
                : clone.getCTTbl().addNewTblPr();
        tPr.set(source.getCTTbl().getTblPr());
        clone.removeRow(0);
        for (XWPFTableRow r : source.getRows()) {
            XWPFTableRow nr = clone.createRow();
            CTTrPr tTrPr = nr.getCtRow().getTrPr() != null ? nr.getCtRow().getTrPr() : nr.getCtRow().addNewTrPr();
            tTrPr.set(r.getCtRow().getTrPr());
            while (nr.getTableCells().size() < r.getTableCells().size()) {
                nr.createCell();
            }

            for (int ic = 0; ic < r.getTableCells().size(); ic += 1) {
                XWPFTableCell c = r.getTableCells().get(ic);
                XWPFTableCell nc = nr.getTableCells().get(ic);

                CTTcPr tTcPr = nc.getCTTc().getTcPr() != null ? nc.getCTTc().getTcPr() : nc.getCTTc().addNewTcPr();
                tTcPr.set(c.getCTTc().getTcPr());

                for (int i = 0; i < c.getParagraphs().size(); i += 1) {
                    XWPFParagraph p = c.getParagraphs().get(i);
                    XWPFParagraph np = nc.getParagraphs().size() > i ? nc.getParagraphs().get(i) : nc.addParagraph();
                    cloneParagraph(np, p);
                }
            }
        }

        return clone;
    }
    
    static final int WAITING = 0, ACTIVE = 1, OPENING = 2, CLOSING = 3, SETTING = 4;
    private static void enumerateParagraph(XWPFParagraph p, int ordinal) {
        /*
         * States:
         *   Waiting - waiting for a tag opening '<<'
         *   Opening - a '<' has been seen, so now we are expecting a second one
         *   Active  - waiting for a 'C' or a '>>'; we are now in a tag string
         *   Closing - a '>' has been seen, so now we are expecting a second one
         *   Setting - a 'C' has been seen, so now we will replace the next 
         *             character with the ordinal if it is an 'N' 
         */

        if (ordinal <= 0)
            return;

        int state = WAITING;
        final LinkedList<XWPFRun> openlist = new LinkedList<>(p.getRuns());
        while (!openlist.isEmpty()) {
            final XWPFRun run = openlist.pop();
            final StringBuilder builder = new StringBuilder(run.text());
            for (int i = 0; i < builder.length(); i += 1) {
                final char c = builder.charAt(i);
                switch (state) {
                    case WAITING: state = c == '<' ? OPENING : WAITING; break;
                    case OPENING: state = c == '<' ? ACTIVE : WAITING; break;
                    case ACTIVE: state = c == '>' ? CLOSING : (c == 'C' ? SETTING : ACTIVE); break;
                    case CLOSING: if (c == '>') state = WAITING; else { state = ACTIVE; i -= 1; } break;
                    case SETTING: 
                        if (c == 'N') {
                            builder.deleteCharAt(i);
                            builder.insert(i, ordinal);
                        }
                        else i -= 1; state = ACTIVE;
                }
            }
            run.setText(builder.toString(), 0);
        }
    }

    private static void enumerateElements(List<IBodyElement> elements, int period) {
        int count = period;
        for (IBodyElement element : elements) {

            int ordinal = count / period;
            count += 1;

            if (element instanceof XWPFParagraph)
                enumerateParagraph((XWPFParagraph) element, ordinal);
            else if (element instanceof XWPFTable) {
                for (XWPFTableRow r : ((XWPFTable) element).getRows()) {
                    for (XWPFTableCell c : r.getTableCells()) {
                        for (XWPFParagraph p : c.getParagraphs()) {
                            enumerateParagraph(p, ordinal);
                        }
                    }
                }
            }
        }
    }

    
    /**
     * Capture all the elements between two specific tags, namely {@code "remove_start"} 
     * and {@code"remove_end"}, and eliminate them from the document. The paragraphs 
     * in which these tags exist will similarly be erased from the document
     * before the it is saved
     * 
     * @param doc   document to be processed
     */
    public static void removeSections(XWPFDocument doc) {
        for (int i = 0, start = 0; i < doc.getBodyElements().size(); i += 1) {
            IBodyElement element = doc.getBodyElements().get(i);
            if (!(element instanceof XWPFParagraph))
                continue;

            final String content = ((XWPFParagraph) element).getText().toLowerCase();
            if (content.contains("remove_start"))
                start = i;

            else if (content.contains("remove_end")) {
                for (int j = i; j >= start; doc.removeBodyElement(j--));
                i = start;
            }
        }
    }
}
