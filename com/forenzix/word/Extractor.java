package com.forenzix.word;

import java.io.PrintWriter;
import java.io.StringWriter;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.LinkedList;
import java.util.List;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFSDT;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.mendix.core.Core;

public class Extractor {
    
    // No object instances for you!
    private Extractor() {
        throw new UnsupportedOperationException("You shall not pass!");
    }

    /*
    // Mendix Specific Code
    @SuppressWarnings({"unchecked", "rawtypes"})
    public static List<Replacer> generateReplacers(XWPFDocument doc, Claim claim, IContext context) {
        final List<String> tags = extractTags(doc);
		final Map<String, Object> vars = new HashMap<>();
		
		final IMendixObject claimCalc = Microflows.get_ClaimCalc(context, claim).getMendixObject();
        final List<Contract> contracts = Core.createXPathQuery("//Calculators.Contract[Calculators.Contract_Claim = $claim]")
            .setVariable("claim", claim.getMendixObject().getId())
            .execute(context).stream()
            .map(m -> Contract.initialize(context, m))
            .collect(Collectors.toList());

		final Map<IMendixIdentifier, IMendixObject> contractCalcs = new HashMap<>();
		
		// Couldn't figure how the hell to write the types for this, so whatever.
		final Set claimMembers = new HashSet<>(claim.getMendixObject().getMembers(context).entrySet());
		claimMembers.addAll(claimCalc.getMembers(context).entrySet());
	
		for (var entry : (Set<Map.Entry<String, IMendixObjectMember<?>>>) claimMembers) {
			final String memName = entry.getKey();
			final Object memVal = entry.getValue().getValue(context);
			vars.put(memName, memVal);
		}
		
		for (int i = 0; i < contracts.size(); i += 1) {
			final Contract contract = contracts.get(i);
			vars.put(String.format("C%d", i + 1), contract.getMendixObject());
			vars.put(String.format("C%d_Ord", i + 1), i + 1);

			contractCalcs.put(contract.getMendixObject().getId(), Microflows.get_ContractCalc(context, contract).getMendixObject());
		}

		// Additional custom variable
		final Account user = usermanagement.proxies.microflows.Microflows.get_User_Account(context);
		if (user != null && user.getFullName() != null) {
			final String name = user.getFullName();

			vars.put("User_Name_FullName", name);
			vars.put("User_Name_Initials", Arrays.asList(name.split("\\s+")).stream()
				.map(s -> s.substring(0, 1))
				.reduce((acc, word) -> acc + "." + word)
				.get().toUpperCase());
		}
		else {
			vars.put("User_Name_FullName", "");
			vars.put("User_Name_Initials", "");
		}

		vars.put("CurrentDateUTC", Date.from(Instant.now()));

		
		final MemberAccessor<Object, String, Object> maccess = (object, member) -> {
			IMendixObject mcalc, mobject = (IMendixObject) object;
			if (contractCalcs.containsKey(mobject.getId()) && (mcalc = contractCalcs.get(mobject.getId())).hasMember(member)) {
				mobject = mcalc;
			}

			final Object value;
			final IMendixObjectMember<?> remember = mobject.getMember(context, member);
			if (remember instanceof MendixEnum) {
				value = Core.getInternationalizedString(
					context, 
					((MendixEnum) remember).getEnumeration()
						.getEnumValues()
						.get(remember.getValue(context))
						.getI18NCaptionKey()
				);
			}
			else {
				value = mobject.getValue(context, member);
			}

			if (!(value instanceof IMendixIdentifier)) return value;
			
			try { 
				return Core.retrieveId(context, (IMendixIdentifier) value); 
			} 
			catch (CoreException ce) { 
				throw new IllegalArgumentException(String.format(
					"Failed to retrieve object with ID %d from Database",
					((IMendixIdentifier) value).toLong()
				)); 
			}
		};
		
		final MemberUpdater<Object, String, Object> mupdate = (object, member, value) -> {
			((IMendixObject) object).setValue(context, member, value instanceof IMendixObject ? 
				((IMendixObject) value).getId() : value
			);
		};

		final List<Replacer> replacers = new LinkedList<>();
		for (String tag : tags) {
			// Assume tags are of the form <<...>> or <<<...>>>
			String prog;
			if (tag.startsWith("<<<") && tag.endsWith(">>>")) {
				// We can some day use triple quoted tags for something. 
				// For now, they are treated the same. -SMG
				prog = tag.substring(3, tag.length() - 3);
			}
			else if (tag.startsWith("<<") && tag.endsWith(">>")) {
				prog = tag.substring(2, tag.length() - 2);
			}
			else {
				throw new IllegalArgumentException("Tag should start and end with double or triple angle brackets (<>)." 
				+ "\nInstead it looks like: " + tag);
			}

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

			// Date and number formats should be dealt with here
			final Interpreter interpreter = new Interpreter(prog, vars);
			interpreter.setMemberAccessCallback(maccess);
			interpreter.setMemberUpdateCallback(mupdate);

			try {
				final Object output = interpreter.interpret();
				replacers.add(replacer(tag, output == null ? "" : format(output, spec)));
			}
			catch (Exception e) {
				report(e, prog);
			}

			// Update our variables so we can have persistance across tag executions.
			vars.putAll(interpreter.getGlobalScopeVariables());
		}

        return replacers;
    }
    */
    
	static Replacer replacer(String bookmark, String replacement) {
		return Replacer.of(bookmark, replacement);
	}

	static void report(Exception e, String program) {
		StringWriter sw = new StringWriter();
		e.printStackTrace(new PrintWriter(sw));

		Core.getLogger("Tag Extractor").error(new StringBuilder("Something went wrong processing the following program:")
			.append("\n").append(program).append("\n").append("Error Details:\n")
			.append(String.format("Type: %s \n", e.getClass().getSimpleName()))
			.append(String.format("Message: %s \n", e.getMessage()))
			.append(String.format("Stack Trace: %s \n", sw.toString())).toString());
	}

	
	private static final DateFormat shortDate = new SimpleDateFormat("dd/MM/yyyy");
 	static String format(Object thing, String spec) {

		if (thing instanceof BigDecimal) return format(((BigDecimal) thing).doubleValue(), spec);
		else if (spec == null) return String.valueOf(thing);
		else if (spec.contains("%")) return String.format(spec, thing);
		else switch (spec) {
			case "short_date":
			case "sdate": return shortDate.format(thing);
			case "currency":
			case "curr": return format(thing, "£ %,.2f");

			// Legacy stuff. I want to phase it out someday.
			case "currp": return String.format("%s p", NumberFormat.getInstance().format((Double) thing * 100));
			case "currpx": return String.format("%s p", NumberFormat.getInstance().format((Double) thing));
			default: throw new IllegalArgumentException("Unrecognised format spec: " + spec);
		}

	}
    

    public static List<String> extractTags(XWPFDocument doc) {
        final StringBuilder builder = new StringBuilder();
        for (IBodyElement element : doc.getBodyElements()) {
            getStringContents(element, builder);
        }

        final String contents = builder.toString();
        final String minimum = "<<>>";
        final List<String> tags = new LinkedList<>();
        for (int i = 0; i < contents.length() - minimum.length() + 1; i += 1) {
            if (contents.charAt(i) == '<' && contents.charAt(i + 1) == '<') {
                for (int j = i + minimum.length() - 1; j < contents.length(); j += 1) {
                    if (contents.charAt(j - 1) == '>' && contents.charAt(j) == '>') {

                        // Handle multiple angle brackets
                        if (j + 1 < contents.length() && contents.charAt(j + 1) == '>') continue;
                        
                        tags.add(contents.substring(i, j + 1));
                        i = j;
                        break;
                    }
                }       
            }
        }
        
        return tags;
    }
    
    public static void getStringContents(IBodyElement element, StringBuilder builder) {
        if (element instanceof XWPFParagraph) {
            getStringContents(((XWPFParagraph) element), builder);
        } else if (element instanceof XWPFTable) {
            getStringContents(((XWPFTable) element), builder);
        } else if (element instanceof XWPFSDT) {
            // Not sure if we want to include Content Tables or not. They're uneditable either way.
            // getStringContents(((XWPFSDT) element), builder);
        } else {
            System.out.println(String.format("Unrecognised IBodyElement type: %s.", element.getClass().getSimpleName()));
        }
    }

    public static void getStringContents(XWPFParagraph paragraph, StringBuilder builder) {
        builder.append(paragraph.getParagraphText());
    }

    public static void getStringContents(XWPFSDT sdt, StringBuilder builder) {
        builder.append(sdt.getContent().getText());
    }
    
    public static void getStringContents(XWPFTable table, StringBuilder builder) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (IBodyElement element : cell.getBodyElements()) {
                    getStringContents(element, builder);
                }
            }
        }
    }
}
