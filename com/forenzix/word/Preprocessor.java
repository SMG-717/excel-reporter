package com.forenzix.word;

import java.util.List;

import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
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
     * Capture all the elements between two specific tags, namely
     * {@code "repeat_start"}
     * and {@code"repeat_end"}, and create multiple copies of them in a row.
     * The paragraphs in which these tags exist will be erased from the document
     * before the document is saved
     * 
     * @param doc   document to be processed
     * @param times the number of times for which the tagged sections shall be
     *              repeated
     */
    public static void repeatSections(XWPFDocument doc, long times) {
        XmlCursor cursor;
        IBodyElement copyable, element;
        int i, j, k, start, amount;

        // Search through the elements of the document
        for (i = 0, start = 0; i < doc.getBodyElements().size(); i += 1) {
            if (!((element = doc.getBodyElements().get(i)) instanceof XWPFParagraph))
                continue;

            // ... until a paragraph that contains the tags 'repeat start' or 'repeat end'
            // is found
            final String content = ((XWPFParagraph) element).getText().toLowerCase();
            if (content.contains("repeat_start"))
                // Record where the repeated section (tags exclusive) starts
                start = i + 1;

            else if (content.contains("repeat_end")) {
                for (k = 1; k < times; k += 1) {
                    // Start copying from the bottom up
                    for (j = i - 1; j >= start; j -= 1) {
                        // Set the cursor right after the current element
                        (cursor = ((XWPFParagraph) element).getCTP().newCursor()).toNextSibling();
                        copyable = doc.getBodyElements().get(j);

                        // And clone the copyable element accordingly
                        if (copyable instanceof XWPFParagraph) {
                            cloneParagraph(doc.insertNewParagraph(cursor), (XWPFParagraph) copyable);
                        } else if (copyable instanceof XWPFTable) {
                            cloneTable(doc.insertNewTbl(cursor), (XWPFTable) copyable);
                        }
                    }
                }

                // Remove the repeat start and end tags, then enumerate all the tags properly
                amount = (i - start) * ((int) times - 1);
                doc.removeBodyElement(i + amount);
                doc.removeBodyElement(start - 1);
                enumerateElements(doc.getBodyElements().subList(start - 1, i + amount), i - start);

                // Advance past the copied section.
                i += amount - 2;
            }
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

    private static void enumerateParagraph(XWPFParagraph p, int ordinal) {

        if (ordinal <= 0)
            return;

        boolean permit = false;
        for (int r = 0; r < p.getRuns().size(); r += 1) {
            XWPFRun run = p.getRuns().get(r);

            // Potential issue
            if (run.text().contains("<<"))
                permit = true;
            else if (run.text().contains(">>"))
                permit = false;

            if (!permit)
                continue;
            if (run.text().contains("CN")) {
                run.setText(run.text().replace("CN", "C" + ordinal), 0);
            }

            // Edge Case: One run ends with C, followed by a run that starts with N.
            if (r < p.getRuns().size() - 1 && run.text().endsWith("C")
                    && p.getRuns().get(r + 1).text().startsWith("N")) {
                p.getRuns().get(r + 1).setText(ordinal + run.text().substring(1), 0);
            }
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
}
