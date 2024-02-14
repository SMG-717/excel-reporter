package com.forenzix.word;

/**
 * The {@code Replacer} class represents a simple structure that does no more
 * than
 * hold a bookmark string, and a replacement string. When a report is generated,
 * the bookmark string is sought out by {@link Replacers} in the document and
 * replaced with replacement string.
 * <p>
 * Replacers are immutable.
 * 
 * @author SMG
 * @see Replacers
 */
public class Replacer {

    public final String bookmark, replacement;

    /**
     * The fields {@code bookmark} and {@code replacement} can be as arbitrarily
     * big as needed, but neither can be null, and bookmark cannot be an empty
     * string.
     * 
     * @param bookmark    A bookmark
     * @param replacement A replacement
     * @throws IllegalArgumentException if either param is null, or if bookmark is
     *                                  blank
     */
    public Replacer(String bookmark, String replacement) {
        if (bookmark == null) {
            throw new IllegalArgumentException("Bookmark cannot be null.");
        } else if (bookmark.isBlank()) {
            throw new IllegalArgumentException("Bookmark cannot be blank.");
        } else if (replacement == null) {
            throw new IllegalArgumentException("Replacement cannot be null.");
        }

        this.bookmark = bookmark;
        this.replacement = replacement;
    }

    /**
     * Returns a simple string representation of this replacer. Tags that span
     * multiple lines are shown as <...>. This function is purely for debugging
     * purposes. Don't use this for anything computational.
     */
    @Override
    public String toString() {
        return (bookmark.contains("\n") ? "<...>" : bookmark) + " -> " + replacement;
    }

    /**
     * Returns a replacer of the supplied bookmark and replacement
     * 
     * @param bookmark    A bookmark string
     * @param replacement A replacement string
     * @return A replacer
     */
    public static Replacer of(String bookmark, String replacement) {
        return new Replacer(bookmark, replacement);
    }
}
