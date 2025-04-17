package com.forenzix.word;

import java.util.LinkedList;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFSDT;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public abstract class Extractor {

    /**
     * Convenience function that extracts the special tags from a document, then
     * generates corresponding replacer objects in one go. This function does not
     * take into account any claim data.
     * <p>
     * Special tags should not include contract specifc operations, because the
     * document has yet to be pre-processed at that stage.
     * 
     * @param doc input document with tags to be extracted
     * @return list of corresponding replacers
     */
    public static List<Replacer> extractSpecialReplacers(XWPFDocument doc) {
        return Replacers.generateReplacers(extractSpecialTags(doc));
    }

    /**
     * Convenience function that extracts tags from a document, then generates
     * corresponding replacer objects in one go. This function does not take into
     * account any claim data.
     * 
     * @param doc input document with tags to be extracted
     * @return list of corresponding replacers
     */
    public static List<Replacer> extractReplacers(XWPFDocument doc) {
        return Replacers.generateReplacers(extractTags(doc));
    }

    /**
     * Extracts all tags from the input document, then filters out any tags which
     * are not special. <b>Special tags</b> start with triple angle brackets.
     * 
     * @param doc input document
     * @return list of special tags
     */
    public static List<String> extractSpecialTags(XWPFDocument doc) {
        return extractTags(doc).stream().filter(s -> s.startsWith("<<<")).collect(Collectors.toList());
    }

    /**
     * Extracts all tags from the input document.
     * <p>
     * <b>Tags</b> are any pieces of text which start with at least 2 opening angle
     * brackets and end with the a matching number of closing brackets. For example,
     * {@code &lt;&lt;&lt;i = i + 1;&gt;&gt;&gt;} is a valid tag, but
     * {@code &lt;&lt;i = i + 1;&gt;&gt;&gt;} is not.
     * <p>
     * Tags can have a format spefication at the end of the tag preceeded by
     * a single colon character. For example {@code &lt;&lt;rate :%.2f&gt;&gt;}
     * will format the given variable as a double with 2 decimal places. For more
     * information about how the format spec works, check out
     * {@link Replacers#format(
     * Object, String) Replacers.format}.
     * 
     * @param doc input document
     * @return list of tags
     */
    public static List<String> extractTags(XWPFDocument doc) {
        // Collect all string content from paragraphs and tables throughout the
        // input document.
        final StringBuilder builder = new StringBuilder();
        for (IBodyElement element : doc.getBodyElements())
            getStringContents(element, builder);

        final String contents = builder.toString();
        final List<String> tags = new LinkedList<>();

        /*
         * The extractor uses a state machine to read and recognise tags form the
         * text content of the input document. The state machine operates as follows:
         * start in a WAITING state
         * WAITING:
         * Look for a series of consequtive less than (<) characters.
         * Once found record their length and change state to BUILDING.
         * 
         * BUILDING:
         * Look for a quote mark (") or greater than (>) character.
         * If a quote mark is found, change stage to STRINGING.
         * If > is found:
         *   Count occurances of > that exist in a row.
         *   If this count matches length:
         *     A valid tag is found and added to the list of tags.
         *     change the state to WAITING.
         * 
         * STRINGING:
         * Look for a quote mark (").
         * If not preceeded by a backward slash (\), change state to BUIDLING.
         */

        // Extractor states
        final int WAITING = 0, // indicates we have yet to encounter a tag
                BUILDING = 1, // we have encountered a tag and currently building it
                STRINGING = 2; // we have encountered a string literal within a tag

        int i, // current character position
                state, // extractor state
                length, // length of current tag opening
                start; // start position of current tag

        for (i = state = length = start = 0; i < contents.length(); i += 1) {
            switch (state) {
                case WAITING:

                    // Count occurances of (<)
                    int j = 0;
                    while (contents.charAt(i + j) == '<')
                        j += 1;

                    if (j <= 1)
                        break;

                    // If we found more than one consequtive (<) characters
                    // start building
                    i = (start = i) + (length = j);
                    state = BUILDING;

                case BUILDING:

                    // Look for (") or (>)
                    switch (contents.charAt(i)) {
                        case '\"':
                            state = STRINGING;
                        default:
                            continue;
                        case '>':
                    }

                    // Count occurances of (>)
                    int k = 0;
                    while (i + k < contents.length() && contents.charAt(i + k) == '>')
                        k += 1;

                    // If we found a matching number of consequtive (>) characters
                    // Collect the tag, and go back to waiting.
                    if (k == length) {
                        tags.add(contents.substring(start, i + k));
                        length = state = WAITING;
                    }

                    // Hence or otherwise, skip the characters we have just processed
                    i += k - 1;
                    break;

                case STRINGING:

                    // Look for a matching quote mark. It has to be one that isn't escaped.
                    if (contents.charAt(i) == '\"' && contents.charAt(i - 1) != '\\')
                        state = BUILDING;
            }
        }

        return tags;
    }

    // Get all text in a given body element
    private static void getStringContents(IBodyElement element, StringBuilder builder) {
        if (element instanceof XWPFParagraph) {
            getStringContents(((XWPFParagraph) element), builder);
        } else if (element instanceof XWPFTable) {
            getStringContents(((XWPFTable) element), builder);
        } else if (element instanceof XWPFSDT) {
            // Not sure if we want to include Content Tables or not. They're uneditable
            // either way.
            // getStringContents(((XWPFSDT) element), builder);
        } else {
            System.out
                    .println(String.format("Unrecognised IBodyElement type: %s.", element.getClass().getSimpleName()));
        }
    }

    // Get all text in a given paragraph
    private static void getStringContents(XWPFParagraph paragraph, StringBuilder builder) {
        builder.append(paragraph.getParagraphText());
    }

    // Get all text in a given table.
    private static void getStringContents(XWPFTable table, StringBuilder builder) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (IBodyElement element : cell.getBodyElements()) {
                    getStringContents(element, builder);
                }
            }
        }
    }
}
