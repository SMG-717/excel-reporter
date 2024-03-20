package com.forenzix.word;

import java.io.PrintWriter;
import java.io.StringWriter;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.AbstractMap.SimpleEntry;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.logging.log4j.core.Core;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFSDT;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import com.forenzix.interpreter.Interpreter;
import com.forenzix.interpreter.Interpreter.MemberAccessor;
import com.forenzix.interpreter.Interpreter.MemberUpdater;

/**
 * This class contains functions for handling {@code Replacer} objects and
 * applying their effects on the given Apache POI objects.
 * 
 * @author SMG
 * @see Replacer
 */
public final class Replacers {

	private static final DateFormat SHORT_DATE_FMT = new SimpleDateFormat("dd/MM/yyyy");
    
    // Mendix Specifc Method
	/**
	 * Extract tags from the input document then use the claim object alond with
	 * its associate objects to find corresponding values for every tag. Contracts
	 * will be included as higher level objects in the interpreter instances and
	 * are aliased 'CN' where N is the index of contract. Note that when higher
	 * level objects are involved, MemberAccess and MemberUpdate handlers need to
	 * be defined for the interpreter to have any functionality for them. 
	 * <p>
	 * Calc object members are included in the interpreter instance as if they're
	 * direct members of their associated parent. For example, claim-level quantum
	 * members have variables that reference them directly, and contract-level
	 * quantum members can be accessed through the associated CN object.
	 * 
	 * @param doc		the input document to generate replacers for
	 * @param claim		the claim object from which data will be sourced
	 * @param context	context object for mendix calls
	 * @return 			list of replacers
	 * @see {@link Interpreter}
	 */
    // @SuppressWarnings({"unchecked", "rawtypes"})
    // public static List<Replacer> generateReplacers(List<String> tags, Claim claim, IContext context) {
	// 	// Initialise the variable map
	// 	final Map<String, Object> vars = new HashMap<>();
		
	// 	// Calculation objects also have their members included in calculation.
	// 	// This could run into an issue if claim/contract and calc objects have
	// 	// members that are identically named.
        
	// 	final IMendixObject claimCalc = Microflows.get_ClaimCalc(context, claim).getMendixObject();
	// 	final Map<IMendixIdentifier, IMendixObject> contractCalcs = new HashMap<>();

    //     final List<Contract> contracts = Core.createXPathQuery("//Calculators.Contract[Calculators.Contract_Claim = $claim]")
    //         .setVariable("claim", claim.getMendixObject().getId())
    //         .execute(context).stream()
    //         .map(m -> Contract.initialize(context, m))
    //         .collect(Collectors.toList());

    //     if (Core.isSubClassOf(calculators.proxies.CCL_BEC_S1.entityName, claim.getMendixObject().getType())) {
    //         final List<BEC_Assumption> assumptions = Core.createXPathQuery("//Calculators.BEC_Assumption")
    //             .execute(context).stream()
    //             .map(m -> BEC_Assumption.initialize(context, m))
    //             .collect(Collectors.toList());

    //         int aCounter = 0;
    //         for (BEC_Assumption a : assumptions) {
    //             final List<IMendixIdentifier> cIds = contracts
    //                 .stream()
    //                 .map(c -> c.getMendixObject().getId())
    //                 .collect(Collectors.toList());

    //             final LinkedList<String> involvedIds = new LinkedList<>();
    //             final Set<IMendixIdentifier> aLinks = ((List<IMendixIdentifier>) a.getMendixObject()
    //                 .getValue(context, "Calculators.Contract_Assumptions"))
    //                 .stream()
    //                 .collect(Collectors.toSet());
    //             for (int i = 0; i < cIds.size(); i += 1) {
    //                 if (aLinks.contains(cIds.get(i))) {
    //                     involvedIds.add(Integer.toString(i + 1));
    //                 }
    //             }

    //             if (involvedIds.size() == 1) {
    //                 vars.put("Assumption_C" + (aCounter += 1), a.getDetails().replace("<<CN>>", involvedIds.getFirst()));
    //             }
    //             else if (!involvedIds.isEmpty()) {
    //                 final String last = involvedIds.removeLast();
    //                 vars.put("Assumption_C" + (aCounter += 1), a.getDetails().replace("<<CN>>", String.join(", ", involvedIds) + " & " + last));
    //             }
    //         }
    //         vars.put("Assumption_Count", aCounter);
    //     }
        

		
	// 	// Couldn't figure how the hell to write the types for this, so whatever.
	// 	final Set claimMembers = new HashSet<>(claim.getMendixObject().getMembers(context).entrySet());
	// 	claimMembers.addAll(claimCalc.getMembers(context).entrySet());
	// 	for (var entry : (Set<Map.Entry<String, IMendixObjectMember<?>>>) claimMembers) {
	// 		final String memName = entry.getKey();
	// 		final Object memVal = entry.getValue().getValue(context);
	// 		vars.put(memName, memVal);
	// 	}
		
	// 	// Contracts are represented by a handful of variables of the form 'CN'.
	// 	// They also get a corresponding Ordinal object for display purposes.
	// 	// Though now that the reporting function uses interpreters and variables
	// 	// can persist through tags, Ord objects are getting more and more obselete.
	// 	for (int i = 0; i < contracts.size(); i += 1) {
	// 		final Contract contract = contracts.get(i);
	// 		vars.put(String.format("C%d", i + 1), contract.getMendixObject());
	// 		vars.put(String.format("C%d_Ord", i + 1), i + 1);

	// 		contractCalcs.put(contract.getMendixObject().getId(), Microflows.get_ContractCalc(context, contract).getMendixObject());
	// 	}

	// 	// Additional custom variables
	// 	final Account user = usermanagement.proxies.microflows.Microflows.get_User_Account(context);
	// 	vars.put("CurrentDateUTC", Date.from(Instant.now()));
	// 	if (user != null && user.getFullName() != null) {
	// 		final String name = user.getFullName();

	// 		vars.put("User_Name_FullName", name);
	// 		vars.put("User_Name_Initials", Arrays.asList(name.split("\\s+")).stream()
	// 			.map(s -> s.substring(0, 1))
	// 			.reduce((acc, word) -> acc + "." + word)
	// 			.get().toUpperCase());
	// 	}
	// 	else {
	// 		vars.put("User_Name_FullName", "");
	// 		vars.put("User_Name_Initials", "");
	// 	}
        
		
	// 	final MemberAccessor<Object, String, Object> maccess = (object, member) -> {
	// 		// When accessing object members, we check first if the member exists
	// 		// in an associated calc object. This calc object is then used instead
	// 		// as the primary object for this operation.
	// 		IMendixObject mcalc, mobject = (IMendixObject) object;
	// 		if (contractCalcs.containsKey(mobject.getId()) && (mcalc = contractCalcs.get(mobject.getId())).hasMember(member)) {
	// 			mobject = mcalc;
	// 		}

	// 		final Object value;
	// 		final IMendixObjectMember<?> remember = mobject.getMember(context, member);

	// 		// If the member is an enum, we want to get its I18N value for display purposes.
	// 		if (remember instanceof MendixEnum) {
	// 			value = Core.getInternationalizedString(
	// 				context, 
	// 				((MendixEnum) remember).getEnumeration()
	// 					.getEnumValues()
	// 					.get(remember.getValue(context))
	// 					.getI18NCaptionKey()
	// 			);
	// 		}
	// 		else {
	// 			value = mobject.getValue(context, member);
	// 		}

	// 		try {
	// 			// If our member points to another object, we grab and return it.
	// 			return value instanceof IMendixIdentifier ? Core.retrieveId(context, (IMendixIdentifier) value) : value;
	// 		} 
	// 		catch (CoreException ce) { 
	// 			throw new IllegalArgumentException(String.format(
	// 				"Failed to retrieve object with ID %d from Database",
	// 				((IMendixIdentifier) value).toLong()
	// 			)); 
	// 		}
	// 	};
		
	// 	// Updating members is direct
	// 	final MemberUpdater<Object, String, Object> mupdate = (object, member, value) -> {
	// 		((IMendixObject) object).setValue(context, member, value instanceof IMendixObject ? 
	// 			((IMendixObject) value).getId() : value
	// 		);
	// 	};

    //     return generateReplacers(tags, maccess, mupdate, vars);
    // }

    
	/**
	 * Generate Replacers from a given list of tags.  
	 * 
	 * @param doc
	 * @param tags
	 * @return
	 */
    public static List<Replacer> generateReplacers(List<String> tags) {
		return generateReplacers(tags, null, null, null);
	}
	
    /**
     * Generate replacers from a given list of tags. MemberAccess and MemberUpdate
     * functions can be provided to instruct the interpreter on how to handle the
     * corresponding events. The set of global variables can be initialised through
     * the inVars parameter. The tags provided must be valid tags and cannot start
     * with any number of (<) characters that is not 2 or 3. Tags can contain
     * format specifications as defined by {@link Replacers#format(Object, String) 
     * Replacers.format}.
     * <p>
     * Note that if an initial variables map inVars is provided, its contents will
     * be altered by the interpreter. This can allow the caller of this function
     * to string together a series of replacers than can all share a continuous
     * interpreter "runtime".
     * 
     * @param tags
     * @param maccess
     * @param mupdate
     * @param inVars
     * @return
     */
    private static List<Replacer> generateReplacers(
		List<String> tags,
		MemberAccessor<Object, String, Object> maccess, 
		MemberUpdater<Object, String, Object> mupdate, 
		Map<String, Object> inVars
	) {

        // If no variable map is provided, we make our own
		final Map<String, Object> vars = inVars == null ? new HashMap<>() : inVars;
		final List<Replacer> replacers = new LinkedList<>();

		for (String tag : tags) {
			String prog;

            // Tags must start with << or <<< and end with >> or >>> respectively
			if (tag.startsWith("<<<") && tag.endsWith(">>>")) {
				prog = tag.substring(3, tag.length() - 3);
			}
			else if (tag.startsWith("<<") && tag.endsWith(">>")) {
				prog = tag.substring(2, tag.length() - 2);
			}
			else {
				throw new IllegalArgumentException("Tag should start and end with double or triple angle brackets (<>)." 
				+ "\nInstead it looks like: " + tag);
			}

            // Find a colon for the format specifier
			final int colon = prog.lastIndexOf(':');
			final int quote = Math.max(prog.lastIndexOf('\"'), prog.lastIndexOf('\''));
			String spec;
			if (colon != -1 && quote < colon) {
				spec = prog.substring(colon + 1, prog.length()).strip();
				prog = prog.substring(0, colon);
			}
			else {
				spec = null;
			}

			final Interpreter interpreter = new Interpreter(prog, vars);
			interpreter.setMemberAccessCallback(maccess);
			interpreter.setMemberUpdateCallback(mupdate);

			try {
                // Core of the Reporter module, where the magic happens.
				final Object output = interpreter.interpret();
				replacers.add(Replacer.of(tag, output == null ? "" : format(output, spec)));
			}
			catch (Exception e) {
				report(e, prog);
			}

			// Update our variables so we can have persistance across tag executions.
			vars.putAll(interpreter.getGlobalScopeVariables());
		}

        return replacers;
    }

    // Log errors to the console
	static void report(Exception e, String program) {
		StringWriter sw = new StringWriter();
		e.printStackTrace(new PrintWriter(sw));

		System.err.println(new StringBuilder("Something went wrong processing the following program:")
			.append("\n").append(program).append("\n").append("Error Details:\n")
			.append(String.format("Type: %s \n", e.getClass().getSimpleName()))
			.append(String.format("Message: %s \n", e.getMessage()))
			.append(String.format("Stack Trace: %s \n", sw.toString())).toString());
	}

    // TODO: document the format spec
 	private static String format(Object thing, String spec) {

		if (thing instanceof BigDecimal) return format(((BigDecimal) thing).doubleValue(), spec);
		else if (spec == null) return String.valueOf(thing);
		else if (spec.contains("%")) return String.format(spec, thing);
		else switch (spec) {
			case "short_date":
			case "sdate": return SHORT_DATE_FMT.format(thing);
			case "currency":
			case "curr": return format(thing, "£ %,.2f");

			// Legacy stuff. I want to phase it out someday.
			case "currp": return String.format("%s p", NumberFormat.getInstance().format((Double) thing * 100));
			case "currpx": return String.format("%s p", NumberFormat.getInstance().format((Double) thing));
			default: throw new IllegalArgumentException("Unrecognised format spec: " + spec);
		}
	}


    /**
     * Traverse through the body elements of the specified document and apply every
     * {@code Replacer} object from the specified list in order. If not all
     * replacers
     * have been applied onto the document, an error is raised.
     * <p>
     * This function relies on the output of {@link Extractor#generateReplacers
     * (XWPFDocument, Claim, IContext) Extractor.generateReplacers}.
     * 
     * @param doc       the target document with replaceable tags
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
     * @param elements  the list of body elements to traverse with replaceable tags
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
                                orderedReplace(cell.getBodyElements(),
                                        replacers.subList(newr.getValue(), replacers.size())))));

                // Update to account for the consumed replacers
                r = newr.getValue();
                e += 1;
            }

            // Otherwise, consume the element
            else
                e += 1;
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
     * <i>Tags</i> are bookmark strings that represent a target for a replacer to
     * plant
     * its replacement string in a paragraph. They can span across one or more runs
     * and/or lines. If the entire tag exists in a single run it could be anywhere
     * within the text of a run.
     * <p>
     * <b>DISCLAIMER</b>: This function is designed to work with Replacers that have been
     * generated by
     * {@link Extractor#generateReplacers(XWPFDocument, Claim, IContext)
     * Extractor.generateReplacers}, otherwise, it might produce unexpected results.
     * 
     * @param p        the specified paragraph where a tag could be found
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
         * return false
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
         * replacement += all characters in f after i
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
     * Generate multiple runs with formatting identical to the original run
     * specified.
     * The number of new runs generated is based on the number of lines in
     * <i>text</i>.
     * If text contains N lines, 2N - 1 runs will be generated, one for each line
     * content, and one for each break {@code &lt;br&gt;} in between. If the
     * provided
     * text is blank or null, no runs will be produced.
     * <p>
     * This function uses {@link Preprocessor#cloneRun(XWPFRun, XWPFRun)
     * Preprocessor.cloneRun}.
     * 
     * @param origin run from which new runs will be cloned
     * @param text   contents of the newly generated runs
     * @throws IllegalArgumentException if the parent of origin isn't a paragraph
     * @see {@link Preprocessor#cloneRun(XWPFRun, XWPFRun) Preprocessor.cloneRun}
     */
    private static void genRuns(XWPFRun origin, String text) {

        // The run has to be a child of a paragraph
        if (!(origin.getParent() instanceof XWPFParagraph)) {
            throw new IllegalArgumentException(
                    String.format("The given origin run is a child of %s which is not a subclass of XWPFParagraph.",
                            origin.getParent().getClass().getName()));
        }

        // Text has to be non-null and non-blank
        if (text == null || text.isBlank())
            return;

        final XWPFParagraph p = (XWPFParagraph) origin.getParent();
        final Iterator<String> itr = text.lines().iterator();
        XWPFRun run;

        // If the string is blank, no runs will be generated.
        if (!itr.hasNext())
            return;

        // Generate a run for the first line of content
        run = Preprocessor.cloneRun(p.insertNewRun(p.getRuns().indexOf(origin) + 1), origin);
        run.setText(itr.next(), 0);
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
