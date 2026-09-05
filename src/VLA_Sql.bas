Attribute VB_Name = "VLA_Sql"
Option Explicit
Public Const VLA_SQL_VERSION As String = "SQL.7"
' SQL.7: WITH CTEs, then recursive WITH - "the convergence proof."
' Two real pieces, sequenced by dependency, not difficulty alone:
'
' (1) Non-recursive WITH: `WITH name AS (query) [, name2 AS (query2)]*
' mainQuery` - each named subresult is a FULL compoundQuery (SQL.6's own
' grammar, unchanged), evaluated in WRITTEN order, and registered as a
' SYNTHETIC table (MakeTableRec) available to every LATER cte AND the
' main query's own FROM/JOIN namespace - the ONLY subquery mechanism
' this engine gets (the roadmap's own words: "a CTE spells every use
' case in this subset with strictly less grammar"). Derived tables
' (`FROM (SELECT ...)`) and correlated subqueries need NO new refusal at
' all: ParseSimpleSelect's own FROM/JOIN grammar already requires a bare
' identifier right there (TokKind <> TK_IDENT fails with the ordinary
' "expected a table name" message), and no scalar/boolean position
' anywhere in this grammar (WHERE/ON/HAVING/a SELECT item) has ever had
' a '(' SELECT ')' case - both are ALREADY structurally impossible,
' confirmed and named here rather than left an unexamined accident.
' A CTE sharing a name with an actually-passed real table SHADOWS it
' (real SQL's own rule) - true for free, not by special-casing: a
' synthetic CTE table record is registered into the SAME `tables`
' Collection SqlRunJoin's own byName dict already builds from, added
' AFTER the real ones, and VLA_Runtime.VlaDictSet is already add-or-
' replace (a later registration under an existing key simply overwrites
' the dict entry) - SqlRunJoin needed zero changes to make this true.
'
' (2) Recursive WITH: `WITH RECURSIVE name AS (base UNION ALL step)
' mainQuery` - the step re-evaluated on the PREVIOUS round's delta only
' (never the full accumulated total) until a round adds zero new rows,
' bounded by a named round ceiling (SQL_CTE_MAX_ROUNDS, matching
' VLA_Datalog's own MAX_ROUNDS value, 10000, exactly) - the roadmap's
' own words, "exactly VLA_Datalog's semi-naive fixpoint, reused (round
' ceiling and all) rather than re-implemented." Examined closely before
' writing any code, not assumed: VLA_Datalog.RunFixpointForRules/
' RunOneRulePass are NOT directly callable here - they are hard-coded
' to Datalog's own body-item/atom/variable representation (EvalRuleBody,
' colOf maps), and VBA has no first-class function value to hand them a
' SQL step query as a generic callback either. What genuinely IS
' reused, literally: the semi-naive ALGORITHM SHAPE itself (round 1 =
' full base evaluation; round N>1 = the step evaluated against ONLY
' round N-1's own delta; stop the moment a round contributes zero new
' rows; a named, loud round-ceiling refusal as the safety valve, never
' a normal exit) - and the mechanism that makes reusing it this cheap is
' SqlRunJoin's own EXISTING table-substitution seam: round N's own
' "delta" is just ANOTHER synthetic MakeTableRec record under the CTE's
' own name, holding ONLY that round's new rows, handed to the SAME
' unchanged SqlRunJoin the step's own branch text already runs through -
' no new evaluation machinery at all, the identical "wrap, don't modify"
' move SQL.6 already made for SqlRunJoin itself. UNION ALL is not a
' stylistic choice here but this item's own grammar requirement (the
' roadmap's own literal words, "base UNION ALL step") - real SQL's own
' rule for a recursive CTE, and the reason no dedup runs across rounds
' at all: two genuinely different rows that happen to compute equal
' values are both kept, exactly as SQLite's own WITH RECURSIVE would
' keep them, with the round ceiling as the ONLY termination safety net
' for a self-referencing step over cyclic data (a real footgun in real
' SQL too, not unique to this engine) - the identical "an in-cell UDF
' has no DoEvents path" reasoning this section's own parent bullet
' already states for PROLOG/SOLVE's own resolution-step budget, applied
' here a second time.
'
' The one recursive SHAPE this item supports, deliberately narrow and
' named rather than generalized: a CTE's own body, when it self-
' references at all, must parse to EXACTLY `ParseCompoundQuery`'s own
' root being a SETOP_UNION_ALL node whose LEFT and RIGHT children are
' BOTH a single COMPOUND_LEAF (never a further compound sub-expression
' on either side) - the base leaf containing NO reference to the CTE's
' own name anywhere in its own FROM/JOIN, the step leaf containing
' EXACTLY ONE. Every other shape is refused by name: a self-reference
' found with no RECURSIVE keyword given at all (sql-cte-self-reference-
' needs-recursive - real SQL's own requirement, most engines including
' SQLite); RECURSIVE given but the body doesn't match the canonical
' base/step shape (sql-cte-recursive-shape - covers a self-reference
' inside an INTERSECT/EXCEPT branch, on the LEFT/base side instead of
' the right, or a chain of more than two top-level branches); the base
' referencing itself (sql-cte-recursive-base-self-references - real
' SQL's own rule, the base case must be non-recursive); the step
' referencing itself MORE than once, e.g. self-joining the recursive CTE
' against itself mid-round (sql-cte-recursive-step-single-reference - a
' real, harder feature, scoped out deliberately, SQL.3's own no-self-
' join precedent repeated); ORDER BY/LIMIT inside a recursive CTE's own
' definition (sql-cte-recursive-no-order-by-limit - per-round ordering
' has no coherent meaning here, unlike a NON-recursive CTE's own body,
' which is an ordinary compoundQuery and keeps ORDER BY/LIMIT for free,
' no new code needed at all). A duplicate CTE name within one WITH
' clause is refused (sql-cte-duplicate-name); an unterminated `(...)`
' CTE definition is refused (sql-cte-unterminated-definition) rather
' than silently consuming the rest of the query text.
'
' Wiring, the same "wrap, don't modify" move repeated a third time:
' SqlRunWith (new, Public) is the actual entry point - queryText with no
' leading WITH keyword at all short-circuits to `SqlRunCompound(
' queryText, tables)` UNCHANGED, byte for byte (proven, not just
' claimed, by VLA_Tests_Query.TestSqlCte), the SAME degenerate-to-the-
' existing-function proof SqlRunCompound itself already gave SqlRunJoin.
' `=SQL(...)` now calls SqlRunWith instead of SqlRunCompound - the ONE
' change to the worksheet-facing entry point, safe for the identical
' reason SQL.6's own SqlRunCompound swap was. Each CTE body's own text
' is sliced out of the ORIGINAL query text by character offset
' (TokStart, SQL.6's own addition, reused unchanged) between a balanced
' `(...)` pair (FindMatchingCloseParen, new - a plain bracket-depth scan
' over TK_OP tokens, no SQL grammar awareness needed at all) - the SAME
' "re-parse real text through the existing pipeline rather than
' evaluate a parsed AST directly" discipline SQL.6's own branch-slicing
' already established.
'
' SQL.6: UNION / UNION ALL / INTERSECT / EXCEPT - a NEW grammar layer,
' not an extension of the existing one. Every SQL.1-SQL.5 item grew ONE
' query's own grammar/pipeline (ParseQuery + SqlRunJoin); this item
' combines MULTIPLE such "simple SELECT" branches with a set operator,
' so it needed its own parsing pass and its own entry point rather than
' another Case inside the existing evaluator.
'
' Grammar (verified against real SQLite behavior, not assumed by
' analogy to a sibling SQL.x item - SQL.5's own header note above is the
' cautionary precedent: its first pass DID reason by analogy from
' GROUP BY's own restriction and was live-caught wrong):
'
'     compoundQuery ::= setOpChain [ORDER BY orderItem [, orderItem]*] [LIMIT NUMBER]
'     setOpChain    ::= intersectChain (('UNION' ['ALL'] | 'EXCEPT') intersectChain)*
'     intersectChain ::= simpleSelect ('INTERSECT' simpleSelect)*
'
' INTERSECT binds TIGHTER than UNION/UNION ALL/EXCEPT, which are left-
' associative with each other - real SQLite precedence, confirmed, not
' guessed - so `A UNION B INTERSECT C EXCEPT D` means
' `(A UNION (B INTERSECT C)) EXCEPT D`. Structurally the SAME "one level
' binds tighter" shape SQL.2's own addExpr-over-mulExpr arithmetic
' precedence already solved: ParseSetOpChain (the outer, UNION/EXCEPT-
' chaining loop) calls ParseIntersectChain (the inner, INTERSECT-
' chaining loop) calls ParseCompoundLeaf (one ordinary simple SELECT) -
' ParseAddExpr/ParseMulExpr's own precedence-climbing shape, mirrored,
' not reinvented. ORDER BY/LIMIT are real, if narrow, additions at the
' COMPOUND level (not scoped out): real SQL allows them only ONCE,
' trailing the LAST branch, applying to the WHOLE combined result, never
' per-branch - ParseCompoundQuery parses them itself, once, after the
' whole setOpChain, the identical block ParseQuery's own SQL.5 grammar
' already parses, factored out into ParseOrderByAndLimit/
' ParseTrailingTerminator so both callers share one grammar rather than
' two copies drifting apart.
'
' ParseSimpleSelect (new, Private) is ParseQuery's own former body,
' extracted verbatim minus its own ORDER BY/LIMIT/trailing-token check -
' a single simple SELECT's grammar boundary, nothing more, reused as
' both ParseQuery's own single-query path (unchanged observable
' behavior - see below) AND ParseCompoundLeaf's own per-branch boundary
' finder. ParseQuery itself is now three calls in sequence
' (ParseSimpleSelect, ParseOrderByAndLimit, ParseTrailingTerminator) -
' a pure internal refactor, byte-for-byte identical parsing behavior to
' before, proven by the fact that VLA_Tests_Query.TestSql/TestSqlJoin
' (every SQL.1-SQL.5 test) needed zero changes.
'
' A compound query's own branch TEXT, not a re-derived AST, is what
' actually gets evaluated: ParseCompoundLeaf records each branch's own
' character-offset SPAN in the ORIGINAL query text (Tokenize gained a
' third token field, TokStart, purely for this - additive only, every
' existing TokKind/TokText call site untouched) and slices it out
' verbatim. EvalCompoundNode then hands that EXACT substring to
' SqlRunJoin, completely UNCHANGED - the "wraps SqlRunJoin, each branch
' runs through the EXISTING pipeline" move the roadmap's own words ask
' for, taken literally rather than approximated by re-deriving a second
' evaluator against a parsed-AST leaf. SqlRunJoin's own signature and
' behavior are therefore untouched by this item, byte for byte - every
' SQL.1-SQL.5 test still calls it directly, unaware SQL.6 exists.
'
' SqlRunCompound (new, Public) is the actual compound entry point, and
' the one place this item's own backward-compatibility promise is made
' real rather than just claimed: when a query has no top-level UNION/
' UNION ALL/INTERSECT/EXCEPT anywhere, ParseCompoundQuery's own tree
' degenerates to a SINGLE leaf, and SqlRunCompound short-circuits to
' `SqlRunJoin(queryText, tables)` - the ORIGINAL text, not the sliced
' branch - so a plain query keeps SQL.5's own ORDER BY SOURCE-column
' fallback intact. That fallback is NOT available for a genuinely
' compound query's own ORDER BY: once two branches are combined there is
' no single coherent source colMap left to fall back to (each branch may
' have joined different tables under different column names), so
' SqlRunCompound resolves ORDER BY items against an EMPTY colMap -
' ResolveOrderItem's own existing SOURCE-fallback branch simply finds
' nothing there and raises the SAME sql-order-by-unknown-column refusal
' a genuinely unresolvable name always gets, rather than a special-cased
' second message - named and tested (VLA_Tests_Query.TestSqlSetOps), not
' left to silently half-work.
'
' Column-count compatibility, not name or type: CombineSetOp checks
' leftArity <> rightArity POST-EVALUATION, never at parse time (a
' SELECT * branch's own arity depends on the actual table it runs
' against, only known once SqlRunJoin has actually run) - refused by
' name, sql-set-op-column-mismatch. The FIRST branch's own output header
' list names the WHOLE compound result regardless of any later branch's
' own aliases - real SQL's own rule, adopted explicitly: CombineSetOp
' always keeps leftResult's own headers, and since "left" is always
' closer to the query's own start at every level of the precedence tree,
' this recursively bubbles the true first branch's headers all the way
' up regardless of tree shape.
'
' Zero new VLA_Relation.bas substrate: UNION ALL is pure concatenation
' (no Relation at all); UNION is dedup-the-concatenation (one RelNew/
' RelTryAdd pass over both sides, in order - SQL.2's own DISTINCT
' precedent, reused directly); INTERSECT is membership (a Relation built
' from the RIGHT side, RelContainsTuple-probed by the LEFT) plus its own
' SEPARATE output-dedup Relation, since real SQL's INTERSECT also
' removes duplicates from its OWN result, not just filters the left
' side; EXCEPT is anti-membership plus that same output-dedup - "not's
' anti-join wearing SQL clothes" (the roadmap's own words), structurally
' the closest of the four to VLA_Datalog.FilterOutMatching's own
' RelContainsTuple-probe shape (read for the SHAPE only - it stays
' DATALOG-private, unification-aware, never hoisted). RelNew/RelTryAdd/
' RelContainsTuple's own existing signatures needed no changes at all.
'
' No parenthesized grouping of compound branches ((A UNION B) INTERSECT
' C) - not asked for by this item, and ParseCompoundLeaf's own first
' token must always be SELECT (ParseSimpleSelect's own opening
' ExpectKeyword), so a leading '(' there already falls through to the
' ordinary "expected SELECT, found '('" refusal - a real, if
' unremarkable, named limit, not a silent gap.
'
' SQL.5: ORDER BY / LIMIT - the "one deliberately-late dissection"
' (BETA_ROADMAP2.md's own words: nothing downstream consumes ordering,
' so this item floats, sequenced here not because anything needs it but
' because it is pure presentation). Genuinely simpler than SQL.4, and
' needs NO new AST node kind and NO new tokenizer work at all - ORDER
' BY/LIMIT items never re-enter the scalar/boolean evaluator (EvalScalar/
' EvalBool gain no new Case at all), they operate on values already
' resolved by the time SqlRunJoin's own row loop runs, the last step of
' its pipeline, downstream of WHERE/GROUP BY/HAVING/SELECT-projection/
' DISTINCT alike (SQL.4's own aggregate pipeline is invisible from here -
' by the time ORDER BY/LIMIT run, accumRows/colMap have already been
' repointed or not, and either way this file just sees "this query's
' own final rows" plus whatever colMap already resolves against them).
'
' ORDER BY item ::= (NUMBER | ident) [ASC | DESC] - resolution (Resolve
' OrderItem) tries TWO namespaces, in order: (1) a 1-based ordinal
' position, or a name matching this query's own OUTPUT header list (a
' SELECT-list item's own alias/bare-column name, fold-compared) - an
' output alias always wins over a same-named source column, real SQL's
' own rule; (2) only when a NAME finds nothing there, it falls back to
' the SAME colMap WHERE/HAVING already resolve bare column names
' against - a real table column (or, for a grouped query, a GROUP BY
' key/aggregate) that was never re-selected at all, real SQL's (and
' specifically SQLite's) own ordinary allowance. This two-namespace
' fallback is NOT the first design tried: v1 resolved ORDER BY against
' the output header list ONLY, reasoning by analogy from SQL.4's own
' GROUP BY restriction (a transformed row's own missing source columns
' have no sound re-evaluation path) - live-caught wrong the moment the
' owner tried the single most ordinary ORDER BY shape there is, `ORDER
' BY dept` over a query that only SELECTed name: refused, where every
' real SQL engine (SQLite included) just sorts by it. The GROUP BY
' analogy doesn't actually transfer here - GROUP BY's own concern was
' that a transformed row's missing columns have NO VALUE AT ALL to sort
' by; ORDER BY's own row loop still has the SOURCE (or grouped) row
' available at the moment it runs, colMap already resolving names
' against it for WHERE/HAVING - so falling back to it costs nothing new,
' only reveals that framing the fallback as "unsound" was the wrong
' abstraction leaked from a different item. What genuinely stays
' restricted, in both design passes: no ARBITRARY expression (`ORDER BY
' salary * 2`), and a SELECT DISTINCT's own dropped columns are still
' unavailable in principle (this engine doesn't special-case DISTINCT
' out of the fallback today - not yet caught live, worth re-examining if
' it ever is). An ordinal out of [1, outArity] or a name matching zero
' or more-than-one candidate (checked in whichever namespace resolves
' it) is refused by name (sql-order-by-position-out-of-range /
' sql-order-by-unknown-column / sql-order-by-ambiguous-column) rather
' than guessed - the ambiguous case mirrors SQL.3's own
' sql-ambiguous-column discipline.
'
' Sorting itself: VBA has no native array sort, and BETA_ROADMAP2.md's
' own scope note is explicit - "one stable sort routine written once...
' over row INDICES, never swapping whole rows per comparison." A merge
' sort (StableSortIndices/MergeSortRange/CompareRowsForSort) over an
' index array - naturally stable by construction (the merge step's own
' tie-break always prefers the LEFT run first), multi-key (walks every
' ORDER BY item in order, falling through to the next key only on an
' exact tie), each key's own comparison reusing VLA_Relation.
' CompareValues directly (CompareOneValue) rather than inventing a
' second value-comparison policy - the identical numeric-vs-text
' STRICT policy WHERE/HAVING already use. DESC simply flips a
' comparison's own sign; ASC (or no direction at all) is the default.
' No second sort routine anywhere else in this codebase needed this yet
' (DATALOG has no ordering concept), so it stays SQL-private for now -
' hoisted into VLA_Relation.bas the moment, not before, a second real
' caller needs the identical mechanism (TableArgResolve's own SQL.1
' precedent, repeated).
'
' LIMIT ::= NUMBER - a plain non-negative INTEGER LITERAL only, never an
' expression (real SQL rarely needs anything else, and "trivial by
' construction" is this item's own roadmap words) - a row-count cutoff
' applied AFTER the sort (or, with no ORDER BY at all, after whatever
' order the query's own evaluation naturally produced - still valid SQL,
' just not a particularly useful one). LIMIT 0 is a real, valid "zero
' rows," not an error.
'
' SQL.4: GROUP BY / aggregates (COUNT/SUM/MIN/MAX/AVG) / HAVING. Three
' new grammar surfaces, none free:
'
' (1) An aggregate CALL - COUNT(...)/SUM(...)/MIN(...)/MAX(...)/AVG(...)
' - is a THIRD thing '(' can mean (alongside a boolean group at
' ParseNotExpr's own level and a scalar group at ParsePrimary's own
' level, SQL.1/SQL.2's own distinction), recognized only when one of
' those five folded identifiers is IMMEDIATELY followed by '(' -
' keyword-in-context the same way SELECT/FROM/WHERE/AS/AND/OR/NOT/
' DISTINCT/INNER/JOIN/ON already are (Tokenize's own header note), so a
' column genuinely named "count" (never followed by '(') still parses
' as an ordinary column reference, unaffected. Parsed inside
' ParsePrimary itself (an aggregate call IS a scalar VALUE, exactly like
' a column or a number) as a new AST node, NK_AGG (NodeAgg/NodeAggKind/
' NodeAggIsStar/NodeAggOperand) - EvalScalar gains exactly one new Case
' for it, never EvalBool (this item's own roadmap words, "zero new
' evaluator code," land on EvalBool specifically; EvalScalar already
' grew a new Case once before, for SQL.2's own NK_ARITH, so this is the
' same move a second time, not a second evaluator). COUNT(*) is the one
' special case - its own operand is '*', not an expression - refused BY
' NAME for every other aggregate (sql-aggregate-star-only-count, real
' SQL's own rule: SUM(*)/MIN(*)/MAX(*)/AVG(*) have no meaning). Nesting
' one aggregate inside another's own operand is refused by name too
' (sql-aggregate-nested, checked via NodeContainsAggregate the moment an
' operand finishes parsing) - real SQL has no meaning for COUNT(SUM(x))
' either. An un-aliased aggregate call is the ONE exception to SQL.2's
' own "a computed SELECT item needs its own AS" rule (ParseQuery's own
' alias-requirement check now also accepts NK_AGG) - COUNT(*) is already
' as self-explanatory a header as a bare column name, unlike an
' arithmetic expression's own ambiguous auto-name; AggDisplayText builds
' that default header (e.g. "COUNT(*)", "SUM(salary)") the one time it's
' needed.
'
' (2) GROUP BY item ::= column (bare or qualified) - deliberately NOT a
' general addExpr, the one real scope narrowing this item makes and
' names explicitly: once grouped, only the GROUP BY key's own values and
' each aggregate's own result survive into the grouped row HAVING/the
' final SELECT list evaluate against (below) - a bare/qualified column
' reference re-resolves there for free (its own colMap key, unchanged),
' but an arbitrary expression's own leaf columns would reference SOURCE
' columns that no longer exist in a grouped row, and re-evaluating it
' against the grouped row has no sound meaning without a second
' evaluator this item does not build. Parsed via ParsePrimary itself
' (reusing its own qualified-column reassembly) with a NodeKind check
' refusing anything else by name (sql-group-by-needs-column) - an
' aggregate call or an arithmetic expression both fail that same check,
' so no separate "aggregate not allowed in GROUP BY" message is needed.
' An aggregate call is also refused, by name, wherever else it would run
' BEFORE grouping exists to summarize anything: WHERE
' (sql-aggregate-not-allowed-in-where) and a JOIN's own ON
' (sql-aggregate-not-allowed-in-on), both checked via the same
' NodeContainsAggregate walk right after each one finishes parsing.
'
' (3) HAVING ::= cond - parsed via ParseOrExpr, the SAME boolean grammar
' WHERE/ON already use (this file's own repeated claim, made real a
' fourth time) - but evaluated against the GROUPED result, never the
' source colMap/row WHERE and ON use, per (4) below.
'
' Whenever a query has ANY aggregate call anywhere in its SELECT list or
' HAVING, or a GROUP BY clause at all (AnySelectItemHasAggregate, all
' parse-time AST shape checks, no colMap/table access needed): SELECT *
' is refused (sql-aggregate-select-star-not-allowed - real SQL's own
' rule, ambiguous which row a bare '*' would even mean); and every OTHER
' SELECT item must be EITHER structurally identical to one of the GROUP
' BY items (ColumnNodesEqual - exact qualifier+name match, no semantic
' column-identity reasoning, a deliberate "refuse rather than guess"
' narrowing named here) OR itself a top-level NK_AGG node
' (sql-select-item-not-grouped otherwise) - real SQL's own "every
' selected column must be a group key or aggregated" rule, stated
' explicitly rather than left implicit (this item's own roadmap words).
'
' (4) The grouping kernel itself is VLA_Relation.RelGroupBy - SQL.4's
' own roadmap text asked for this by name ("whatever shared grouping
' kernel falls out of unifying [DATALOG.2's count/sum, SQL.2's DISTINCT,
' and this item] belongs in VLA_Relation.bas, not the SQL module");
' VLA_Relation.bas's own new header section has the full design.
' ComputeGroupedRows (below) is this file's own bridge into it:
' enumerates every DISTINCT aggregate call across the SELECT list and
' HAVING (AggSignature's own dedup - the same call named twice shares
' one computed slot), builds one "eval row" per WHERE-surviving joined
' row (its own GROUP BY key values, then each NEEDED aggregate operand's
' own value - COUNT needs none at all, VLA_Relation.bas's own header
' note has the reason), and calls RelGroupBy. The result is one row per
' group; a matching colMap (groupColMap - keyed exactly the way
' AddColumnsToMap already keys a real table's own columns for every
' GROUP BY column, plus each aggregate's own canonical AggSignature text
' for its own result column) is built alongside it. SqlRunJoin then
' simply REPOINTS accumRows/colMap at the grouped result/its own colMap
' and hasWhere/whereExpr at hasHaving/havingExpr - every line of code
' below that point (SELECT-list header building, the row-filter-then-
' project-then-DISTINCT loop) runs COMPLETELY UNCHANGED against
' whichever pipeline actually ran, this item's own "zero new evaluator
' code, zero new projection code" payoff made real, not just claimed.
' HAVING may reference a raw, non-grouped source column only if it's
' ALSO a GROUP BY key (real SQL's own rule) - never any other source
' column - because groupColMap holds nothing else; a HAVING column
' reference outside the GROUP BY key set fails ResolveColumnPos's own
' unchanged sql-unknown-column check, the identical refusal an ordinary
' unknown column already gets, no new message needed for this narrower
' case at all.
'
' MIN/MAX/AVG over zero rows (only reachable for an ungrouped aggregate
' query - SELECT COUNT(*)/... FROM t WHERE <matches nothing>, no GROUP
' BY at all - a real GROUP BY group can never be empty by construction)
' has no sensible numeric default and is refused by name
' (sql-aggregate-empty-no-rows), rather than inventing a NULL sentinel
' this engine has no representation for anywhere else (VLA_Relation.bas'
' own header note has the full reasoning); COUNT/SUM over the same case
' are real, meaningful zeros, DATALOG.2's own precedent, unchanged.
' DISTINCT after a grouped/aggregated SELECT needed no decision at all -
' it already operates on the FINAL PROJECTED output row regardless of
' where accumRows came from, so a grouped query gets it for free, the
' identical mechanism SQL.2 already built.
'
' SQL.3: INNER JOIN ... ON, multiple tables folded left-to-right. Grammar
' grows one clause - FROM tablename (INNER JOIN | JOIN) tablename ON
' cond, zero or more times, "INNER" optional and meaning nothing extra
' since a bare JOIN already means INNER JOIN in real SQL (SQLite
' included) - parsed by ParseQuery exactly as before, ON reusing
' ParseOrExpr, the SAME boolean grammar WHERE already uses (one
' evaluator/one grammar, pointed at a third place). LEFT/RIGHT/FULL/
' OUTER/CROSS are refused BY NAME (sql-outer-join-not-supported),
' checked at both positions a JOIN keyword could legally start, rather
' than falling through to the generic end-of-query leftover-token
' check - BETA_ROADMAP2.md's own words, "a dissection, not a
' microscope," get an actual teaching refusal, not just a vague one.
'
' Reuses VLA_Relation.RelJoin's own hash join DIRECTLY, never a second
' nested-loop implementation (BETA_ROADMAP1.md's own standing law) -
' the one real substrate gap this forced open: RelJoin's own accessors
' (RelArity/RelTuples/RelCount) expect a Relation-SHAPED 3-item record,
' but SQL's own rows are a BAG (RangeToRows, never RelFromRange/
' RelTryAdd, SQL.1's own load-bearing distinction) - wrapping a bag via
' RelNew+RelTryAdd to feed RelJoin would silently DEDUP two identical
' source rows, dropping a legitimate join match. VLA_Relation.RelWrapBag
' (new, Public) sidesteps this: the SAME 3-item shape, Item(3)=Nothing
' instead of a real membership index, safe specifically because RelJoin
' itself never reads Item(3) at all (only RelTryAdd/RelContainsTuple do,
' neither ever called on a join's own input sides here) - its own header
' note in VLA_Relation.bas has the full safety argument. VLA_Datalog.bas's
' own former Private CollToLongArray also hoisted there (Public now) the
' moment this item needed the identical Collection-of-Longs -> RelJoin's
' own leftCols()/rightCols() conversion for equi-condition join keys -
' both the same second-consumer move TableArgResolve/RangeColumnNames/
' CompareValues already made.
'
' The parser splits each ON into equi-conditions (SplitOnExpr, walking
' only through NK_AND at the top: a bare column = bare column comparison
' where one side resolves onto the already-accumulated side and the
' other onto the newly-joined table becomes ONE hash-join key; multiple
' such terms compose into a genuine COMPOSITE key, RelJoin's own
' leftCols()/rightCols() already being arrays, not scalars) plus a
' residual filter (everything else - an OR, a NOT, a non-equi comparison,
' a column-vs-constant comparison, two columns from the SAME side -
' folded into one AND-chain, evaluated once per joined row through
' EvalBool, ZERO new evaluation code, this file's own SQL.1 sequencing
' note made real a third time). Applied incrementally, right after each
' join step, not deferred to one filter at the very end - for a pure
' INNER JOIN chain the two orderings produce the identical final row set
' (filtering never adds rows), so this is a genuine performance choice,
' not a correctness one, and the more natural translation of "this ON
' clause constrains only this join."
'
' Qualified column names (Employees.id) enter the name-resolution map
' here, since two joined tables can otherwise collide on a bare column
' name - the tokenizer gained '.' as a new single-char operator (never
' reached inside a number literal, whose own digit-scan already claims
' any internal '.' first), ParsePrimary's own TK_IDENT case reassembles
' IDENT '.' IDENT into one NK_COLUMN node carrying an OPTIONAL fourth
' item (qualifierFolded, "" for every SQL.1/SQL.2 call site and every
' unqualified reference), and colMap (AddColumnsToMap, ResolveColumnPos
' - the one new chokepoint EvalScalar/ValidateScalarColumns/SplitOnExpr
' all now call instead of a bare VlaDictHas/Get pair) holds BOTH a
' qualified key ("tablename.colname", always precise) and an unqualified
' one for every column NOT already claimed by an earlier table - a
' second table registering the SAME bare name doesn't overwrite the
' first table's own entry, it marks it Empty, a sentinel ResolveColumnPos
' turns into a named refusal (sql-ambiguous-column) rather than a wrong
' answer. A single-table query (SQL.1/SQL.2, or a SQL.3 query with no
' JOIN at all) never exercises the collision branch - one table can
' never collide with itself - so this degenerates to exactly the old
' flat colMap in that case, unchanged.
'
' No table aliasing (AS after a table name) - not asked for by this
' item, and real SQL needs it specifically for a self-join, which this
' item also does not support: the SAME table name appearing twice across
' FROM/JOIN is refused by name (sql-duplicate-join-table) rather than
' left to silently misresolve every one of its own column references
' between two indistinguishable occurrences. SqlRun's own ORIGINAL
' single-table signature is UNCHANGED - every SQL.1/SQL.2 pure test still
' calls it directly - now a thin wrapper around the generalized
' SqlRunJoin (a one-table Collection), the identical "extend without
' breaking" move VLA_Datalog.DatalogRun's own Optional baseRelations/
' headerMap already made twice. LEFT JOIN deliberately deferred
' (BETA_ROADMAP2.md's own words: "a dissection, not a microscope" -
' DATALOG's own `not` already proves the substrate can express the
' anti-join half; nothing downstream needs it yet).
'
' SQL.2: computed SELECT expressions, AS aliases, DISTINCT. The scalar
' grammar grows a real arithmetic hierarchy - addExpr (+/-, left-
' associative) over mulExpr (*//, left-associative) over primary
' (column/number/string, unary +/-, or a parenthesized addExpr) - and
' ParseComparison's own operands now call ParseAddExpr instead of
' ParsePrimary directly, so WHERE gets arithmetic operands "for free"
' (salary * 1.1 > 100000 works with no parens) rather than needing a
' second grammar. NK_ARITH joins the AST (NodeArith/NodeArithOp/
' NodeArithLeft/NodeArithRight, NodeCmp's own shape repeated); unary
' minus is (0 - operand), reusing the binary "-" evaluator path rather
' than inventing a unary-specific one.
'
' The SELECT list itself is no longer a bare column-name list - each
' item is now addExpr [AS ident] (MakeProjItem/ProjItemExpr/
' ProjItemHasAlias/ProjItemAlias), and DISTINCT is a keyword checked
' right after SELECT. A computed (non-bare-column) item WITHOUT its own
' AS alias is refused at parse time (sql-computed-column-needs-alias) -
' a deliberate scoping call, not a limitation discovered later: real SQL
' auto-names an unaliased computed column from its own expression text,
' but this engine's own "refuse rather than silently guess" culture
' argues against inventing that name-from-text convention when GROUP
' BY/ORDER BY will need to refer back to it reliably later (this item's
' own roadmap words - "AS matters beyond convenience"). NodeColumn grows
' a THIRD item (the column's own as-typed original text, alongside its
' already-folded name) purely so a bare, un-aliased SELECT column can
' still header itself with its own original casing, the same as SQL.1
' already did before the SELECT list stopped being raw strings.
'
' DISTINCT's own whole-row dedup reuses VLA_Relation.RelNew/RelTryAdd
' directly - a Relation already IS "a case-sensitive SET of Variant
' tuples, Mac-fallback aware," exactly what DISTINCT needs, so no new
' substrate or even a new Public surface was needed in VLA_Relation.bas
' for this (unlike DATALOG.4's own ValueIsNumericType/CompareValues/
' ComputeArithmetic hoist) - RelNew/RelTryAdd were already Public.
' ComputeArithmetic (DATALOG.4's own hoist) gets its SECOND real caller
' here (EvalScalar's own new NK_ARITH case), proving that hoist out:
' SQL's own numeric-ness policy stays STRICT (ValueIsNumericType alone,
' unchanged from EvalBool's own NK_CMP case) - never DATALOG's own
' lenient numeric-looking-string policy, since SQL's values are always
' real Excel Value2 reads.
'
' The one deliberate, DOCUMENTED limit: a comparison's own LEFT operand
' can never itself START with '(' at the very top of a condition (WHERE
' (a + b) > 10 is refused, not silently misparsed) - ParseNotExpr's own
' '(' check (SQL.1's own live-caught fix, unchanged) always claims a
' leading '(' there as an attempted boolean group FIRST, before
' ParseComparison ever gets a chance to try parsing a parenthesized
' scalar instead; resolving that ambiguity for real would need bounded
' lookahead past the matching ')' - real machinery, not "grammar only,"
' so it stays out of THIS item's scope. Every other position (a
' comparison's own RIGHT operand, anything nested inside an arithmetic
' expression once already committed to scalar context, and any SELECT-
' list item, which never touches ParseNotExpr at all) is unaffected and
' fully supports parenthesized grouping.
'
' =====================================================================
'  VLA_Sql - the QUERY AND LOGIC section's second engine, following the
'  split BETA_ROADMAP2.md's own SQL item names: VLA_Relation.bas (the
'  shared, DSL-ignorant substrate DATALOG already proved out) plus this
'  thin module owning only SQL's own grammar and its own =SQL(...) UDF.
'
'  SQL.1/SQL.2 (this file's own scope so far): single-table SELECT/
'  WHERE, computed SELECT expressions, AS aliases, DISTINCT. Real SQL
'  text, not a fourth S-expression dialect (the section's own frozen
'  decision, 2026-08-29 - "when a canonical surface syntax already
'  exists in the intended users' heads, adopt it and pay the parser
'  cost"): a real tokenizer plus recursive-descent parser over a NAMED,
'  FROZEN subset -
'
'      SELECT [DISTINCT] projItem [, projItem]* | [DISTINCT] *
'      FROM tablename joinClause*
'      [WHERE cond]
'      [;]
'
'      joinClause ::= ('INNER' 'JOIN' | 'JOIN') tablename 'ON' cond
'      projItem ::= addExpr [AS ident]
'      cond ::= cond OR cond | cond AND cond | NOT cond | '(' cond ')'
'             | addExpr (= | <> | < | <= | > | >=) addExpr
'      addExpr ::= mulExpr (('+' | '-') mulExpr)*
'      mulExpr ::= primary (('*' | '/') primary)*
'      primary ::= column | number | 'string' | '(' addExpr ')'
'             | '-' primary | '+' primary
'      column ::= ident | ident '.' ident
'
'      SQL.3: joinClause's own ON reuses cond, the SAME grammar WHERE
'      already uses - one evaluator, pointed at a third place, this
'      file's own repeated claim made real again. LEFT/RIGHT/FULL/
'      OUTER/CROSS JOIN are refused BY NAME (sql-outer-join-not-
'      supported), not part of this frozen subset yet.
'
'      '(' cond ')' (a boolean group) and '(' addExpr ')' (a scalar
'      group) are DIFFERENT grammar rules that happen to share a
'      bracket character, reachable from DIFFERENT positions - cond's
'      own top level (ParseNotExpr) versus deep inside a comparison's
'      operand or a SELECT-list item (ParsePrimary). Conflating them
'      broke exactly this in an earlier pass: NOT (col = 'x') parsed
'      the whole parenthesized group as if it were ParseComparison's
'      own left OPERAND, then went hunting for a comparison operator
'      after a group that had already completed one internally.
'      ParseNotExpr's own header has the full post-mortem, and this
'      file's own SQL.2 header note above has the one remaining
'      documented limit this split still leaves: a comparison's own
'      LEFT operand can never itself START with '(' at the very top of
'      a condition (ParseNotExpr always claims that spot for a boolean
'      group first).
'
'      SQL.6: the FULL query text this subset describes may itself be
'      one BRANCH of a compound query -
'
'          compoundQuery ::= setOpChain [ORDER BY orderItem [, ...]] [LIMIT NUMBER]
'          setOpChain    ::= intersectChain (('UNION' ['ALL'] | 'EXCEPT') intersectChain)*
'          intersectChain ::= <the query text above> ('INTERSECT' <the query text above>)*
'
'      INTERSECT binds tighter than UNION/UNION ALL/EXCEPT, which are
'      left-associative with each other (real SQLite precedence,
'      verified). ORDER BY/LIMIT trail the WHOLE compound query exactly
'      once, after the last branch, never per-branch. No parenthesized
'      grouping of branches - VLA_Sql.bas's own SQL.6 header note has
'      the full design and every named refusal (sql-set-op-column-
'      mismatch and kin).
'
'      SQL.7: the WHOLE query text this subset describes may itself be
'      preceded by named subresults -
'
'          topLevelQuery ::= [withClause] compoundQuery
'          withClause    ::= 'WITH' ['RECURSIVE'] cteDef (',' cteDef)*
'          cteDef        ::= ident 'AS' '(' compoundQuery ')'
'
'      Each cteDef's own name joins the FROM/JOIN namespace for every
'      LATER cteDef and the final compoundQuery, exactly like a real
'      passed table - the ONLY subquery mechanism this subset gets;
'      derived tables and correlated subqueries are already structurally
'      impossible in this grammar (no position anywhere ever accepts a
'      '(' SELECT ')'), confirmed and named rather than left an
'      unexamined accident. A cteDef whose own compoundQuery references
'      its own name needs the RECURSIVE keyword and must match ONE
'      shape - `base UNION ALL step`, neither side itself a further
'      compound expression - VLA_Sql.bas's own SQL.7 header note has the
'      full design and every named refusal (sql-cte-recursive-shape and
'      kin).
'
'  Everything outside this subset is refused by name at parse time
'  (sql-unsupported-keyword and kin), never guessed - datalog-compound-
'  term's own discipline, applied to a second grammar. Query text is
'  locale-invariant: number literals always use '.' as the decimal
'  separator (ParseInvariantNumber's own header has the reason), string
'  literals are single-quoted with '' as an escaped quote (the ordinary
'  SQL convention), regardless of the caller's own Excel locale - SD-4's
'  freeze, decided before any engine code, BETA_ROADMAP2.md's own words.
'
'  One SQLite-shaped dialect, deliberately, not a cherry-picked mix
'  (BETA_ROADMAP2.md's own dialect-pin note: mixing dialects invalidates
'  all of them de facto, and the whole point of real SQL text is that a
'  query copied from an actual tool just works) - '' string-escaping,
'  <> not !=, case-sensitive string DATA comparison (vbBinaryCompare,
'  never a silent NOCASE), case-INSENSITIVE identifier matching
'  (VLA_Identity.Fold, SQLite's own default for unquoted names) all
'  checked against real SQLite behavior, not assumed. Both of SQLite's
'  own comment forms are recognized and skipped like whitespace ('--' to
'  end of line, '/* ... */'), and exactly one trailing ';' is tolerated
'  - both ordinary artifacts of a query actually pasted from a real
'  tool, the scenario this whole grammar-surface decision exists to
'  serve, not just a hand-typed one.
'
'  SQL.1's own two microscopes every later SQL.x item consumes rather
'  than re-deriving: the parser skeleton (its refuse-by-name discipline,
'  its token/AST shapes) and the scalar expression evaluator (EvalScalar/
'  EvalBool) - SQL.2's own computed columns (below) are this SAME
'  evaluator pointed at the projection list instead of only the filter,
'  proving the claim out; SQL.3's residual ON conditions and SQL.4's
'  HAVING are meant to be the same move again, never a second evaluator.
'
'  Named-column resolution is new substrate DATALOG never needed (it is
'  deliberately positional) - VLA_Relation.RangeColumnNames, column
'  names folded once through VLA_Identity.Fold (SD-8's law, never a
'  second LCase - the neutrality audit's own Turkish-I warning already
'  flags that trap). Matched rows collect into BufAdd's own manually-
'  doubled buffer, never per-row ReDim Preserve (the O(n^2) trap this
'  item exists to kill, named because SQL's own motivation - scripts/
'  spreadsheet.lisp's SELECTROWS crashing on large tables - is a real,
'  stated performance requirement DATALOG never had to answer to).
'
'  Stated MVP limits, named rather than silently missing: at least ONE
'  table argument (sql-needs-at-least-one-table - SQL.1's own now-retired
'  sql-needs-exactly-one-table, never reused, SD-9); every table argument
'  must be a real Excel Table, ListObject-backed (sql-table-needs-real-
'  table) - a plain named range has no column names of its own for SQL
'  to select/filter by, unlike DATALOG's own positional convention which
'  never needed any; every table FROM/JOIN names must correspond to one
'  of the passed arguments (sql-from-table-mismatch/sql-join-table-not-
'  passed), matched by NAME never by argument position, the same
'  auditability discipline DATALOG's own table-predicate-naming already
'  established; the same table name may not appear twice across FROM/
'  JOIN (sql-duplicate-join-table - no aliasing, so a self-join has no
'  way to disambiguate its own two occurrences); only INNER JOIN (or a
'  bare JOIN) is built, LEFT/RIGHT/FULL/OUTER/CROSS refused by name
'  (sql-outer-join-not-supported); an unqualified column existing in
'  more than one joined table is refused (sql-ambiguous-column), not
'  guessed toward whichever table registered it first; a computed SELECT
'  item without its own AS alias is refused (sql-computed-column-needs-
'  alias, SQL.2's own scoping call, this file's own SQL.2 header note
'  above has the reasoning); a comparison's own LEFT operand can never
'  itself START with '(' at the very top of a condition (this file's own
'  SQL.2 header note above has the full mechanism).
'
'  LAYER:     1 (VLA_Relation)
'  MAY CALL:  VLA_Relation (RangeToRows, RangeColumnNames,
'             TableArgResolve, ValueIsNumericType, CompareValues,
'             ComputeArithmetic, InvariantVal, RelNew, RelTryAdd,
'             RelContainsTuple - SQL.6's own first caller, VLA_Datalog's
'             own stratified-negation precedent, reused unchanged - RelJoin,
'             RelWrapBag, CollToLongArray - SQL.3's own two new
'             calls, both hoisted from an existing single-caller module
'             the moment a second one needed the identical mechanism -
'             SQL.2's own DISTINCT reuses RelNew/RelTryAdd directly as a
'             ready-made case-sensitive whole-row dedup set, no new
'             substrate needed - the whole substrate this engine reuses
'             rather than re-deriving), VLA_Identity (Fold), VLA_Runtime
'             (VlaDictNew/Set/Get/Has - column-name-to-position maps,
'             identifier lookups only, never row data), VLA_Messages
'             (RaiseMsg).
'  SHIPS:     add-in only, as a real Excel worksheet function
'             (=SQL(...)) the moment the add-in loads - no VBProject
'             trust, no export, no emitted code. SEC.1 Tier-0 by
'             construction: reads only the ranges/text it is given,
'             writes nothing, calls nothing outside itself.
'  PAYS INTO: QUERY AND LOGIC's own SQL item; the shared-substrate bet
'             VLA_Relation.bas exists to prove, DATALOG's own precedent
'             reused rather than re-derived for a second engine.
' =====================================================================

' ---- token kinds -----------------------------------------------------
Private Const TK_IDENT As Long = 0
Private Const TK_NUMBER As Long = 1
Private Const TK_STRING As Long = 2
Private Const TK_OP As Long = 3
Private Const TK_EOF As Long = 4

' ---- AST node kinds ----------------------------------------------------
Private Const NK_COLUMN As Long = 0
Private Const NK_NUMBER As Long = 1
Private Const NK_STRING As Long = 2
Private Const NK_CMP As Long = 3
Private Const NK_AND As Long = 4
Private Const NK_OR As Long = 5
Private Const NK_NOT As Long = 6
Private Const NK_ARITH As Long = 7
Private Const NK_AGG As Long = 8

' ---- SQL.5: ORDER BY item kinds - a 1-based ordinal position into the
'      output header list, or a fold-compared name matching one of
'      those headers (this file's own SQL.5 header note has the full
'      resolution rule) ---------------------------------------------
Private Const ORDER_ORDINAL As Long = 0
Private Const ORDER_NAME As Long = 1

' ---- SQL.5: where a resolved ORDER BY item's own sort VALUE comes
'      from - the already-projected OUTPUT row, or the SOURCE/grouped
'      row colMap resolves everything else against (ResolveOrderItem's
'      own header note has the full "why", below) -------------------
Private Const ORDERSRC_OUTPUT As Long = 0
Private Const ORDERSRC_SOURCE As Long = 1

' ---- SQL.6: compound-query tree node kinds - a LEAF holds one already-
'      sliced branch's own query TEXT (evaluated by re-running it
'      through the existing, unchanged SqlRunJoin); an internal node
'      holds one of the three set operators plus its own left/right
'      subtree (MakeCompoundSetOp's own header note has the shape) ----
Private Const COMPOUND_LEAF As Long = 0
Private Const SETOP_UNION As Long = 1
Private Const SETOP_UNION_ALL As Long = 2
Private Const SETOP_INTERSECT As Long = 3
Private Const SETOP_EXCEPT As Long = 4

' ---- SQL.7: a recursive WITH's own round ceiling - the SAME numeric
'      value as VLA_Datalog's own MAX_ROUNDS (10000), a deliberate match
'      rather than an independently-chosen number, even though the two
'      loops share no code (VBA has no first-class function value to
'      hand a generic round-runner, and Datalog's own fixpoint is hard-
'      coded to its own atom/body-item representation) - this file's own
'      SQL.7 header note has the full reasoning. A named, loud safety
'      valve for a self-referencing step over cyclic data, never a
'      normal exit. -----------------------------------------------------
Private Const SQL_CTE_MAX_ROUNDS As Long = 10000

' ---- token records: Item(1)=kind (Long), Item(2)=text (String),
'      Item(3)=startPos (Long, 1-based character offset into the
'      ORIGINAL query text this token began at - SQL.6's own addition,
'      purely so ParseCompoundLeaf can slice a branch's own ORIGINAL
'      text back out of a single shared tokenization; every existing
'      TokKind/TokText call site is untouched, additive only) ---------
Private Function MakeToken(ByVal kind As Long, ByVal text As String, ByVal startPos As Long) As Collection
    Dim t As New Collection
    t.Add kind
    t.Add text
    t.Add startPos
    Set MakeToken = t
End Function

Private Function TokKind(ByVal t As Collection) As Long
    TokKind = t.Item(1)
End Function

Private Function TokText(ByVal t As Collection) As String
    TokText = t.Item(2)
End Function

Private Function TokStart(ByVal t As Collection) As Long
    TokStart = t.Item(3)
End Function

Private Function IsDigitChar(ByVal c As String) As Boolean
    IsDigitChar = (c >= "0" And c <= "9")
End Function

Private Function IsIdentStartChar(ByVal c As String) As Boolean
    IsIdentStartChar = (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or c = "_"
End Function

Private Function IsIdentChar(ByVal c As String) As Boolean
    IsIdentChar = IsIdentStartChar(c) Or IsDigitChar(c)
End Function

' The earliest CR or LF at or after fromPos, or 0 if the text has
' neither (a '--' comment then runs to the end of the string) - checked
' as two separate InStr calls rather than one, since a query built via
' TEXTJOIN(CHAR(10), ...) (DATALOG's own multi-line convention) may use
' bare LF with no CR at all.
Private Function FindLineEnd(ByVal s As String, ByVal fromPos As Long) As Long
    Dim eolR As Long, eolN As Long
    eolR = InStr(fromPos, s, vbCr)
    eolN = InStr(fromPos, s, vbLf)
    If eolR = 0 Then
        FindLineEnd = eolN
    ElseIf eolN = 0 Then
        FindLineEnd = eolR
    ElseIf eolR < eolN Then
        FindLineEnd = eolR
    Else
        FindLineEnd = eolN
    End If
End Function

' Breaks queryText into tokens. Keywords (SELECT/FROM/WHERE/AND/OR/NOT)
' are NOT their own token kind - they are plain TK_IDENT tokens the
' PARSER recognizes by folded text (VLA_Identity.Fold, StrComp), the
' same "reserved word in context, not lexically" shape SQL itself uses.
' A single-quoted string uses '' as an escaped quote, the ordinary SQL
' convention (never a backslash escape, which real SQL doesn't use
' either). Any character outside this frozen subset - SQL.2 added +, -,
' / to the single-char operator set below (joining SELECT *'s own *) -
' is refused here, by name, rather than silently accepted and misparsed
' three layers up.
Private Function Tokenize(ByVal s As String) As Collection
    Dim toks As New Collection
    Dim i As Long, n As Long
    i = 1
    n = Len(s)
    Do While i <= n
        Dim c As String
        c = Mid$(s, i, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then
            i = i + 1
        ElseIf c = "'" Then
            Dim j As Long, buf As String
            j = i + 1
            buf = ""
            Dim closed As Boolean
            closed = False
            Do While j <= n
                If Mid$(s, j, 1) = "'" Then
                    If j < n Then
                        If Mid$(s, j + 1, 1) = "'" Then
                            buf = buf & "'"
                            j = j + 2
                        Else
                            closed = True
                            Exit Do
                        End If
                    Else
                        closed = True
                        Exit Do
                    End If
                Else
                    buf = buf & Mid$(s, j, 1)
                    j = j + 1
                End If
            Loop
            If Not closed Then VLA_Messages.RaiseMsg "sql-unterminated-string"
            toks.Add MakeToken(TK_STRING, buf, i)
            i = j + 1
        ElseIf IsDigitChar(c) Then
            Dim k As Long, numBuf As String, dotSeen As Boolean
            k = i
            numBuf = ""
            dotSeen = False
            Do While k <= n
                Dim ck As String
                ck = Mid$(s, k, 1)
                If IsDigitChar(ck) Then
                    numBuf = numBuf & ck
                    k = k + 1
                ElseIf ck = "." And Not dotSeen Then
                    dotSeen = True
                    numBuf = numBuf & ck
                    k = k + 1
                Else
                    Exit Do
                End If
            Loop
            toks.Add MakeToken(TK_NUMBER, numBuf, i)
            i = k
        ElseIf IsIdentStartChar(c) Then
            Dim m As Long, idBuf As String
            m = i
            idBuf = ""
            Do While m <= n
                If IsIdentChar(Mid$(s, m, 1)) Then
                    idBuf = idBuf & Mid$(s, m, 1)
                    m = m + 1
                Else
                    Exit Do
                End If
            Loop
            toks.Add MakeToken(TK_IDENT, idBuf, i)
            i = m
        ElseIf c = "<" Then
            If i < n And Mid$(s, i + 1, 1) = "=" Then
                toks.Add MakeToken(TK_OP, "<=", i)
                i = i + 2
            ElseIf i < n And Mid$(s, i + 1, 1) = ">" Then
                toks.Add MakeToken(TK_OP, "<>", i)
                i = i + 2
            Else
                toks.Add MakeToken(TK_OP, "<", i)
                i = i + 1
            End If
        ElseIf c = ">" Then
            If i < n And Mid$(s, i + 1, 1) = "=" Then
                toks.Add MakeToken(TK_OP, ">=", i)
                i = i + 2
            Else
                toks.Add MakeToken(TK_OP, ">", i)
                i = i + 1
            End If
        ElseIf c = "-" And i < n And Mid$(s, i + 1, 1) = "-" Then
            ' A '--' line comment - real SQLite syntax, ordinary in any
            ' query pasted from an actual SQL tool (BETA_ROADMAP2.md's
            ' own SQLite-dialect pin exists precisely so pasted text
            ' works, not just hand-typed text). Checked BEFORE the bare
            ' '-' arithmetic-operator branch below, so '--' still always
            ' wins the race - only a LONE '-' (no second '-' immediately
            ' after) ever reaches that branch.
            Dim eol As Long
            eol = FindLineEnd(s, i)
            If eol = 0 Then
                i = n + 1
            Else
                i = eol + 1
            End If
        ElseIf c = "/" And i < n And Mid$(s, i + 1, 1) = "*" Then
            ' A '/* ... */' block comment - the other real SQLite/ANSI
            ' comment form. Unterminated is refused, not silently
            ' extended to end-of-string - the same "never silently
            ' guess" discipline as an unterminated string literal.
            ' Checked BEFORE the bare '/' arithmetic-operator branch
            ' below, the identical "comment wins the race" ordering as
            ' '--' above - only a LONE '/' (no '*' immediately after)
            ' ever reaches that branch.
            Dim endPos As Long
            endPos = InStr(i + 2, s, "*/")
            If endPos = 0 Then VLA_Messages.RaiseMsg "sql-unterminated-comment"
            i = endPos + 2
        ElseIf c = "=" Or c = "," Or c = "(" Or c = ")" Or c = "*" Or c = ";" Or c = "+" Or c = "-" Or c = "/" Or c = "." Then
            ' ';' tokenizes unconditionally (any real SQL text may
            ' contain one); ParseQuery's own leftover-token check is
            ' what decides where exactly one is tolerated (immediately
            ' before end-of-query, the ordinary "one pasted statement,
            ' optionally semicolon-terminated" case) versus refused (a
            ' second statement after it - SQL.1 answers one query, not
            ' a script). SQL.2: +, -, / join the single-char operator
            ' set - a bare '-' or '/' reaching here (past the comment
            ' checks above, which always run first) is unambiguously
            ' arithmetic, never a comment starter. SQL.3: '.' joins the
            ' set too, for a qualified column reference (Employees.id) -
            ' only ever reached OUTSIDE a number literal, since a digit-
            ' led numeric token's own scan (IsDigitChar branch, above)
            ' already consumes any '.' immediately following digits as
            ' part of that number first; an identifier's own scan
            ' (IsIdentChar) never includes '.', so a qualified reference
            ' always tokenizes as IDENT '.' IDENT, three tokens, which
            ' ParsePrimary's own TK_IDENT case (below) reassembles.
            toks.Add MakeToken(TK_OP, c, i)
            i = i + 1
        Else
            VLA_Messages.RaiseMsg "sql-unexpected-character", "char", c
        End If
    Loop
    toks.Add MakeToken(TK_EOF, "", i)
    Set Tokenize = toks
End Function

' Val() is genuinely LOCALE-INVARIANT in VBA - it always treats '.' as
' the decimal separator, unlike CDbl/CSng, which respect the user's own
' regional settings. This is exactly the '.' -always requirement
' BETA_ROADMAP2.md's own SQL item freezes "decided now rather than
' inherited by accident" - CDbl("50000.5") would silently misparse as
' 500005 on a comma-decimal Excel locale, exactly the trap this function
' exists to avoid. The tokenizer's own numBuf construction already
' guarantees at most one '.', so Val's own behavior here is unambiguous.
' DATALOG.4: delegates to VLA_Relation.InvariantVal, the identical
' Val()-wraps-locale-invariance reasoning hoisted the moment a second
' engine needed it too - unchanged behavior, one fewer duplicate.
Private Function ParseInvariantNumber(ByVal s As String) As Double
    ParseInvariantNumber = VLA_Relation.InvariantVal(s)
End Function

' ---- AST node records: Item(1)=kind (Long); shape past that depends
'      on kind, exactly like VLA_Datalog.bas's own body-item records ---
Private Function NodeKind(ByVal n As Collection) As Long
    NodeKind = n.Item(1)
End Function

' SQL.2: gained a THIRD item, originalText - the column's own as-typed
' casing (e.g. "Salary"), alongside its already-folded name (e.g.
' "salary") - purely so a bare, un-aliased SELECT-list column can still
' header its own output with its own original casing, the same as
' SQL.1 already did back when the SELECT list was raw strings instead
' of AST nodes. WHERE's own NK_COLUMN usage (EvalScalar/EvalBool) never
' reads Item(3) at all - unaffected, additive-only change.
' SQL.3: gained a FOURTH item, qualifierFolded - a table qualifier's own
' folded name (Employees.id's own "employees"), "" when the reference is
' unqualified (every SQL.1/SQL.2 call site, and every single-table
' SQL.3 query too) - Optional so no existing call site needed to change.
Private Function NodeColumn(ByVal foldedName As String, ByVal originalText As String, Optional ByVal qualifierFolded As String = "") As Collection
    Dim n As New Collection
    n.Add NK_COLUMN
    n.Add foldedName
    n.Add originalText
    n.Add qualifierFolded
    Set NodeColumn = n
End Function

Private Function NodeColumnName(ByVal n As Collection) As String
    NodeColumnName = n.Item(2)
End Function

Private Function NodeColumnOriginal(ByVal n As Collection) As String
    NodeColumnOriginal = n.Item(3)
End Function

Private Function NodeColumnQualifier(ByVal n As Collection) As String
    NodeColumnQualifier = n.Item(4)
End Function

Private Function NodeNumber(ByVal v As Double) As Collection
    Dim n As New Collection
    n.Add NK_NUMBER
    n.Add v
    Set NodeNumber = n
End Function

Private Function NodeNumberValue(ByVal n As Collection) As Double
    NodeNumberValue = n.Item(2)
End Function

Private Function NodeString(ByVal v As String) As Collection
    Dim n As New Collection
    n.Add NK_STRING
    n.Add v
    Set NodeString = n
End Function

Private Function NodeStringValue(ByVal n As Collection) As String
    NodeStringValue = n.Item(2)
End Function

Private Function NodeCmp(ByVal op As String, ByVal l As Collection, ByVal r As Collection) As Collection
    Dim n As New Collection
    n.Add NK_CMP
    n.Add op
    n.Add l
    n.Add r
    Set NodeCmp = n
End Function

Private Function NodeCmpOp(ByVal n As Collection) As String
    NodeCmpOp = n.Item(2)
End Function

Private Function NodeCmpLeft(ByVal n As Collection) As Collection
    Set NodeCmpLeft = n.Item(3)
End Function

Private Function NodeCmpRight(ByVal n As Collection) As Collection
    Set NodeCmpRight = n.Item(4)
End Function

' SQL.2: NodeCmp's own shape repeated for arithmetic - op is one of
' +/-/*// (unary minus/plus never reach here, ParsePrimary's own header
' has that mechanism - they build straight off NK_ARITH's binary "-"
' via a literal 0 operand instead of a fifth op string).
Private Function NodeArith(ByVal op As String, ByVal l As Collection, ByVal r As Collection) As Collection
    Dim n As New Collection
    n.Add NK_ARITH
    n.Add op
    n.Add l
    n.Add r
    Set NodeArith = n
End Function

Private Function NodeArithOp(ByVal n As Collection) As String
    NodeArithOp = n.Item(2)
End Function

Private Function NodeArithLeft(ByVal n As Collection) As Collection
    Set NodeArithLeft = n.Item(3)
End Function

Private Function NodeArithRight(ByVal n As Collection) As Collection
    Set NodeArithRight = n.Item(4)
End Function

' SQL.4: an aggregate call - kind is one of VLA_Relation.AGG_COUNT/
' AGG_SUM/AGG_MIN/AGG_MAX/AGG_AVG, isStar True only for COUNT(*)
' (operand then Nothing and never read), operand an addExpr AST node for
' every other case.
Private Function NodeAgg(ByVal kind As Long, ByVal isStar As Boolean, ByVal operand As Collection) As Collection
    Dim n As New Collection
    n.Add NK_AGG
    n.Add kind
    n.Add isStar
    n.Add operand
    Set NodeAgg = n
End Function

Private Function NodeAggKind(ByVal n As Collection) As Long
    NodeAggKind = n.Item(2)
End Function

Private Function NodeAggIsStar(ByVal n As Collection) As Boolean
    NodeAggIsStar = n.Item(3)
End Function

Private Function NodeAggOperand(ByVal n As Collection) As Collection
    Set NodeAggOperand = n.Item(4)
End Function

Private Function NodeBin(ByVal kind As Long, ByVal l As Collection, ByVal r As Collection) As Collection
    Dim n As New Collection
    n.Add kind
    n.Add l
    n.Add r
    Set NodeBin = n
End Function

Private Function NodeBinLeft(ByVal n As Collection) As Collection
    Set NodeBinLeft = n.Item(2)
End Function

Private Function NodeBinRight(ByVal n As Collection) As Collection
    Set NodeBinRight = n.Item(3)
End Function

Private Function NodeNot(ByVal operand As Collection) As Collection
    Dim n As New Collection
    n.Add NK_NOT
    n.Add operand
    Set NodeNot = n
End Function

Private Function NodeNotOperand(ByVal n As Collection) As Collection
    Set NodeNotOperand = n.Item(2)
End Function

' ---- recursive-descent parser -----------------------------------------
Private Function CurTok(ByVal toks As Collection, ByVal pos As Long) As Collection
    Set CurTok = toks.Item(pos)
End Function

Private Function DescribeTok(ByVal t As Collection) As String
    If TokKind(t) = TK_EOF Then
        DescribeTok = "end of query"
    Else
        DescribeTok = "'" & TokText(t) & "'"
    End If
End Function

' kw must already be lowercase - VLA_Identity.Fold lowercases ASCII, so
' comparing a folded token's text against a lowercase literal is the
' correct, SD-8-consistent way to recognize a keyword regardless of the
' case the query text itself used.
Private Function AtKeyword(ByVal toks As Collection, ByVal pos As Long, ByVal kw As String) As Boolean
    Dim t As Collection
    Set t = CurTok(toks, pos)
    AtKeyword = (TokKind(t) = TK_IDENT) And (StrComp(VLA_Identity.Fold(TokText(t)), kw, vbBinaryCompare) = 0)
End Function

Private Function AtOp(ByVal toks As Collection, ByVal pos As Long, ByVal op As String) As Boolean
    Dim t As Collection
    Set t = CurTok(toks, pos)
    AtOp = (TokKind(t) = TK_OP) And (TokText(t) = op)
End Function

Private Sub ExpectKeyword(ByVal toks As Collection, ByRef pos As Long, ByVal kw As String)
    If Not AtKeyword(toks, pos, kw) Then
        VLA_Messages.RaiseMsg "sql-expected-token", "expected", UCase$(kw), "found", DescribeTok(CurTok(toks, pos))
    End If
    pos = pos + 1
End Sub

Private Sub ExpectOp(ByVal toks As Collection, ByRef pos As Long, ByVal op As String)
    If Not AtOp(toks, pos, op) Then
        VLA_Messages.RaiseMsg "sql-expected-token", "expected", "'" & op & "'", "found", DescribeTok(CurTok(toks, pos))
    End If
    pos = pos + 1
End Sub

' ---- SELECT-list projection-item records (SQL.2): Item(1)=exprNode
'      (Collection, an addExpr AST node - a bare column reference is
'      just NK_COLUMN, no special-casing needed anywhere downstream),
'      Item(2)=hasAlias (Boolean), Item(3)=aliasName (String, "" and
'      unused when hasAlias is False) ------------------------------
Private Function MakeProjItem(ByVal exprNode As Collection, ByVal hasAlias As Boolean, ByVal aliasName As String) As Collection
    Dim p As New Collection
    p.Add exprNode
    p.Add hasAlias
    p.Add aliasName
    Set MakeProjItem = p
End Function

Private Function ProjItemExpr(ByVal p As Collection) As Collection
    Set ProjItemExpr = p.Item(1)
End Function

Private Function ProjItemHasAlias(ByVal p As Collection) As Boolean
    ProjItemHasAlias = p.Item(2)
End Function

Private Function ProjItemAlias(ByVal p As Collection) As String
    ProjItemAlias = p.Item(3)
End Function

' ---- JOIN-clause records (SQL.3): Item(1)=tableName (String, folded -
'      SqlRunJoin's own byName lookup resolves it against the actually-
'      passed table arguments, exactly as fromName already does), Item(2)
'      =onExpr (Collection, an already-parsed ParseOrExpr AST - the SAME
'      boolean grammar WHERE already uses, reused rather than growing a
'      second one). Splitting onExpr into equi-conditions (fed to
'      RelJoin as join keys) vs. a residual post-join filter happens
'      LATER, in SqlRunJoin (SplitOnExpr) - it needs the ACTUAL passed
'      tables' own column layout to know which side of the join a given
'      column belongs to, which ParseQuery (pure grammar, no execution
'      context, this file's own header note) has no way to know. ------
Private Function MakeJoinRec(ByVal tableName As String, ByVal onExpr As Collection) As Collection
    Dim r As New Collection
    r.Add tableName
    r.Add onExpr
    Set MakeJoinRec = r
End Function

Private Function JoinRecTable(ByVal r As Collection) As String
    JoinRecTable = r.Item(1)
End Function

Private Function JoinRecOn(ByVal r As Collection) As Collection
    Set JoinRecOn = r.Item(2)
End Function

' ---- ORDER BY item records (SQL.5): Item(1)=kind (ORDER_ORDINAL or
'      ORDER_NAME), Item(2)=ordinalPos (Long, unused/0 for ORDER_NAME),
'      Item(3)=nameFolded (String, unused/"" for ORDER_ORDINAL),
'      Item(4)=descending (Boolean) - resolution against the actual
'      output header list happens later, in SqlRunJoin, once outHeaders
'      is known (ParseQuery has no execution context, its own header
'      note, repeated a third time) -----------------------------------
Private Function MakeOrderItem(ByVal kind As Long, ByVal ordinalPos As Long, ByVal nameFolded As String, ByVal descending As Boolean) As Collection
    Dim r As New Collection
    r.Add kind
    r.Add ordinalPos
    r.Add nameFolded
    r.Add descending
    Set MakeOrderItem = r
End Function

Private Function OrderItemKind(ByVal r As Collection) As Long
    OrderItemKind = r.Item(1)
End Function

Private Function OrderItemOrdinal(ByVal r As Collection) As Long
    OrderItemOrdinal = r.Item(2)
End Function

Private Function OrderItemName(ByVal r As Collection) As String
    OrderItemName = r.Item(3)
End Function

Private Function OrderItemDescending(ByVal r As Collection) As Boolean
    OrderItemDescending = r.Item(4)
End Function

' SQL.3's own explicit, named refusal for every JOIN variant this item
' does NOT build (BETA_ROADMAP2.md's own words: "LEFT JOIN deliberately
' deferred... it is a dissection, not a microscope") - checked wherever
' a JOIN keyword could legally start (right after FROM's own table, and
' again after each completed join clause), so LEFT/RIGHT/FULL/OUTER/
' CROSS get this SPECIFIC, teaching refusal rather than falling through
' to ParseQuery's own generic end-of-query "leftover token" check, which
' would still refuse it but without explaining that only INNER JOIN (or
' a bare JOIN) is built so far.
Private Sub RefuseUnsupportedJoinKind(ByVal toks As Collection, ByVal pos As Long)
    Dim kw As Variant
    For Each kw In Array("left", "right", "full", "outer", "cross")
        If AtKeyword(toks, pos, CStr(kw)) Then
            VLA_Messages.RaiseMsg "sql-outer-join-not-supported", "keyword", UCase$(CStr(kw))
        End If
    Next kw
End Sub

' SQL.6: the single-simple-SELECT grammar - SELECT/DISTINCT/FROM/JOIN/
' WHERE/GROUP BY/HAVING, stopping wherever that grammar naturally ends
' (a top-level ORDER BY/LIMIT/UNION/INTERSECT/EXCEPT keyword, ';', or
' EOF) WITHOUT consuming or checking any of that - former ParseQuery's
' own body, extracted verbatim (byte-for-byte identical logic, only the
' Tokenize call and everything from ORDER BY onward moved out) so it can
' be reused as BOTH ParseQuery's own single-query path (unchanged
' observable behavior, proven by every SQL.1-SQL.5 test needing zero
' changes) AND ParseCompoundLeaf's own per-branch boundary finder
' (SQL.6) - "how far does one simple SELECT's own grammar extend in this
' shared token stream" is the ONE thing both callers need, nothing more.
Private Sub ParseSimpleSelect(ByVal toks As Collection, ByRef pos As Long, ByRef hasDistinct As Boolean, ByRef selectAll As Boolean, ByRef selectItems As Collection, _
                        ByRef fromName As String, ByRef joins As Collection, ByRef whereExpr As Collection, ByRef hasWhere As Boolean, _
                        ByRef hasGroupBy As Boolean, ByRef groupByItems As Collection, ByRef hasHaving As Boolean, ByRef havingExpr As Collection)
    ExpectKeyword toks, pos, "select"

    ' SQL.2: DISTINCT is checked right after SELECT, before either *
    ' or the projection list - it applies to the WHOLE output row
    ' either way (SqlRun's own whole-row dedup, below), never per-item.
    hasDistinct = False
    If AtKeyword(toks, pos, "distinct") Then
        hasDistinct = True
        pos = pos + 1
    End If

    selectAll = False
    Set selectItems = New Collection
    If AtOp(toks, pos, "*") Then
        selectAll = True
        pos = pos + 1
    Else
        Do
            ' SQL.2: each item is now addExpr [AS ident], not a bare
            ' column-name token - a bare column is simply the NK_COLUMN
            ' case of that same grammar, so EvalScalar handles every
            ' item uniformly at row-projection time (SqlRun, below),
            ' no separate "is this a plain column" branch needed there.
            Dim exprNode As Collection
            Set exprNode = ParseAddExpr(toks, pos)

            Dim hasAlias As Boolean, aliasName As String
            hasAlias = False
            aliasName = ""
            If AtKeyword(toks, pos, "as") Then
                pos = pos + 1
                Dim at As Collection
                Set at = CurTok(toks, pos)
                If TokKind(at) <> TK_IDENT Then
                    VLA_Messages.RaiseMsg "sql-expected-token", "expected", "an alias name", "found", DescribeTok(at)
                End If
                aliasName = TokText(at)
                hasAlias = True
                pos = pos + 1
            End If

            ' A computed (non-bare-column) item without its own alias
            ' is refused here, at parse time - this file's own SQL.2
            ' header note has the scoping reasoning (sql-computed-
            ' column-needs-alias). SQL.4: an un-aliased AGGREGATE call
            ' is exempted from this rule - unlike an arithmetic
            ' expression's own ambiguous auto-name, an aggregate call's
            ' own canonical text (COUNT(*), SUM(salary)) is already as
            ' self-explanatory a header as a bare column name
            ' (AggDisplayText, used when SqlRunJoin builds this item's
            ' own output header below).
            If Not hasAlias And NodeKind(exprNode) <> NK_COLUMN And NodeKind(exprNode) <> NK_AGG Then
                VLA_Messages.RaiseMsg "sql-computed-column-needs-alias"
            End If

            selectItems.Add MakeProjItem(exprNode, hasAlias, aliasName)
            If AtOp(toks, pos, ",") Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
        ' No "selectItems.Count = 0" check here - genuinely unreachable
        ' given the loop's own shape: it either raises (a bad token, or
        ' a missing alias) on the very first item or adds at least one
        ' item before ever checking for a comma, so an empty list can
        ' never survive to be counted - SQL.1's own identical reasoning,
        ' unchanged.
    End If

    ExpectKeyword toks, pos, "from"

    Dim ft As Collection
    Set ft = CurTok(toks, pos)
    If TokKind(ft) <> TK_IDENT Then
        VLA_Messages.RaiseMsg "sql-expected-token", "expected", "a table name", "found", DescribeTok(ft)
    End If
    fromName = VLA_Identity.Fold(TokText(ft))
    pos = pos + 1

    ' SQL.3: (INNER JOIN | JOIN) tablename ON cond, zero or more, folded
    ' left-to-right (SqlRunJoin's own job, not this parser's - this loop
    ' only ever builds the WRITTEN sequence of join records). "INNER" is
    ' optional and means nothing extra - a bare JOIN already means INNER
    ' JOIN in real SQL (SQLite included), so both spellings are accepted
    ' as the identical thing, the same "adopt what a pasted query already
    ' says" stance the whole grammar surface is built on. Every OTHER
    ' join kind (LEFT/RIGHT/FULL/OUTER/CROSS) is refused by name, not
    ' silently misread as INNER - checked at both positions a JOIN
    ' keyword could legally start: right here, and again after each
    ' completed join clause below.
    Set joins = New Collection
    RefuseUnsupportedJoinKind toks, pos
    Do While AtKeyword(toks, pos, "inner") Or AtKeyword(toks, pos, "join")
        If AtKeyword(toks, pos, "inner") Then
            pos = pos + 1
            ExpectKeyword toks, pos, "join"
        Else
            pos = pos + 1
        End If
        Dim jt As Collection
        Set jt = CurTok(toks, pos)
        If TokKind(jt) <> TK_IDENT Then
            VLA_Messages.RaiseMsg "sql-expected-token", "expected", "a table name", "found", DescribeTok(jt)
        End If
        Dim joinTableName As String
        joinTableName = VLA_Identity.Fold(TokText(jt))
        pos = pos + 1
        ExpectKeyword toks, pos, "on"
        Dim onExpr As Collection
        Set onExpr = ParseOrExpr(toks, pos)
        ' SQL.4: an aggregate can never appear in ON - ON filters rows
        ' before grouping ever happens (this file's own SQL.4 header
        ' note).
        If NodeContainsAggregate(onExpr) Then VLA_Messages.RaiseMsg "sql-aggregate-not-allowed-in-on"
        joins.Add MakeJoinRec(joinTableName, onExpr)
        RefuseUnsupportedJoinKind toks, pos
    Loop

    hasWhere = False
    Set whereExpr = Nothing
    If AtKeyword(toks, pos, "where") Then
        pos = pos + 1
        hasWhere = True
        Set whereExpr = ParseOrExpr(toks, pos)
        ' SQL.4: an aggregate can never appear in WHERE either, the
        ' identical reason as ON above - filters rows before grouping
        ' ever happens.
        If NodeContainsAggregate(whereExpr) Then VLA_Messages.RaiseMsg "sql-aggregate-not-allowed-in-where"
    End If

    ' SQL.4: GROUP BY item ::= column (bare or qualified) - deliberately
    ' NOT a general addExpr, this file's own SQL.4 header note has the
    ' full scope reasoning. Reuses ParsePrimary itself (its own
    ' qualified-column reassembly, SQL.3) with a NodeKind check refusing
    ' anything else by name (sql-group-by-needs-column) - an aggregate
    ' call or an arithmetic expression both fail this same check, so no
    ' separate "aggregate not allowed in GROUP BY" message is needed.
    hasGroupBy = False
    Set groupByItems = New Collection
    If AtKeyword(toks, pos, "group") Then
        pos = pos + 1
        ExpectKeyword toks, pos, "by"
        hasGroupBy = True
        Do
            Dim gExpr As Collection
            Set gExpr = ParsePrimary(toks, pos)
            If NodeKind(gExpr) <> NK_COLUMN Then
                VLA_Messages.RaiseMsg "sql-group-by-needs-column"
            End If
            groupByItems.Add gExpr
            If AtOp(toks, pos, ",") Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
    End If

    ' SQL.4: HAVING ::= cond, the SAME boolean grammar WHERE/ON already
    ' use (this file's own repeated claim, made real a fourth time) -
    ' unlike WHERE/ON, an aggregate call IS allowed here (the entire
    ' point of HAVING), so no NodeContainsAggregate refusal runs against
    ' it.
    hasHaving = False
    Set havingExpr = Nothing
    If AtKeyword(toks, pos, "having") Then
        pos = pos + 1
        hasHaving = True
        Set havingExpr = ParseOrExpr(toks, pos)
    End If

    ' SQL.4: whenever this query aggregates at all (a GROUP BY clause,
    ' a HAVING clause, or an aggregate call anywhere in the SELECT
    ' list), SELECT * is refused (real SQL's own rule - ambiguous which
    ' row a bare '*' would even mean once rows have collapsed into
    ' groups) and every OTHER SELECT item must be either structurally
    ' identical to one of the GROUP BY items or itself a top-level
    ' aggregate call - this file's own SQL.4 header note has the full
    ' reasoning. All parse-time AST shape checks, no colMap/table access
    ' needed at all.
    If hasGroupBy Or hasHaving Or AnySelectItemHasAggregate(selectItems) Then
        If selectAll Then VLA_Messages.RaiseMsg "sql-aggregate-select-star-not-allowed"
        Dim svi As Long
        For svi = 1 To selectItems.Count
            Dim sExpr As Collection
            Set sExpr = ProjItemExpr(selectItems.Item(svi))
            If NodeKind(sExpr) <> NK_AGG Then
                Dim matched As Boolean
                matched = False
                Dim sgi As Long
                For sgi = 1 To groupByItems.Count
                    If ColumnNodesEqual(sExpr, groupByItems.Item(sgi)) Then
                        matched = True
                        Exit For
                    End If
                Next sgi
                If Not matched Then
                    VLA_Messages.RaiseMsg "sql-select-item-not-grouped", "position", svi
                End If
            End If
        Next svi
    End If

End Sub

' SQL.5: ORDER BY item ::= (NUMBER | ident) [ASC | DESC] - resolved
' against the OUTPUT header list later, in SqlRunJoin/SqlRunCompound
' (this file's own SQL.5 header note has the full reasoning); ASC/DESC
' are checked the SAME keyword-in-context way as every other reserved
' word here (AtKeyword against folded text) - no tokenizer change
' needed at all, unlike SQL.4's own new aggregate-call grammar.
'
' SQL.6: extracted out of former ParseQuery so ParseCompoundQuery's own
' outer grammar (ORDER BY/LIMIT trail the WHOLE compound query exactly
' once, after the LAST branch, never per-branch) can share this SAME
' block instead of a second copy drifting apart from it - ParseQuery
' below is now the FIRST caller of this shared helper, not the only one.
Private Sub ParseOrderByAndLimit(ByVal toks As Collection, ByRef pos As Long, _
                                  ByRef hasOrderBy As Boolean, ByRef orderItems As Collection, _
                                  ByRef hasLimit As Boolean, ByRef limitCount As Long)
    hasOrderBy = False
    Set orderItems = New Collection
    If AtKeyword(toks, pos, "order") Then
        pos = pos + 1
        ExpectKeyword toks, pos, "by"
        hasOrderBy = True
        Do
            Dim ot As Collection
            Set ot = CurTok(toks, pos)
            Dim oKind As Long, oOrdinal As Long, oName As String
            oOrdinal = 0
            oName = ""
            If TokKind(ot) = TK_NUMBER Then
                oKind = ORDER_ORDINAL
                Dim ordVal As Double
                ordVal = ParseInvariantNumber(TokText(ot))
                If ordVal < 1 Or ordVal <> Int(ordVal) Then
                    VLA_Messages.RaiseMsg "sql-order-by-invalid-position", "value", TokText(ot)
                End If
                oOrdinal = CLng(ordVal)
                pos = pos + 1
            ElseIf TokKind(ot) = TK_IDENT Then
                oKind = ORDER_NAME
                oName = VLA_Identity.Fold(TokText(ot))
                pos = pos + 1
            Else
                VLA_Messages.RaiseMsg "sql-expected-token", "expected", "a column name or output position", "found", DescribeTok(ot)
            End If
            Dim oDesc As Boolean
            oDesc = False
            If AtKeyword(toks, pos, "asc") Then
                pos = pos + 1
            ElseIf AtKeyword(toks, pos, "desc") Then
                oDesc = True
                pos = pos + 1
            End If
            orderItems.Add MakeOrderItem(oKind, oOrdinal, oName, oDesc)
            If AtOp(toks, pos, ",") Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
    End If

    ' SQL.5: LIMIT ::= NUMBER - a plain non-negative integer LITERAL
    ' only, never an expression (this file's own SQL.5 header note has
    ' the reasoning); applied as a row-count cutoff after ORDER BY (or,
    ' with no ORDER BY, after whatever order the query's own evaluation
    ' naturally produced), entirely in SqlRunJoin/SqlRunCompound, once
    ' the final row set is known.
    hasLimit = False
    limitCount = 0
    If AtKeyword(toks, pos, "limit") Then
        pos = pos + 1
        Dim lt As Collection
        Set lt = CurTok(toks, pos)
        If TokKind(lt) <> TK_NUMBER Then
            VLA_Messages.RaiseMsg "sql-expected-token", "expected", "a non-negative whole number", "found", DescribeTok(lt)
        End If
        Dim limVal As Double
        limVal = ParseInvariantNumber(TokText(lt))
        If limVal < 0 Or limVal <> Int(limVal) Then
            VLA_Messages.RaiseMsg "sql-limit-needs-nonneg-integer", "value", TokText(lt)
        End If
        hasLimit = True
        limitCount = CLng(limVal)
        pos = pos + 1
    End If
End Sub

' Exactly one trailing ';' is tolerated (the ordinary "one pasted
' statement, semicolon-terminated" shape any real SQL tool produces) -
' anything AFTER it is a second statement, which this engine answers one
' query at a time and refuses, not silently truncates. Any OTHER
' leftover token (a keyword this subset doesn't support, or - for
' ParseQuery's own single-query caller - a top-level UNION/INTERSECT/
' EXCEPT that only ParseCompoundQuery's own grammar handles) is refused
' by name (sql-unsupported-keyword) rather than silently ignored.
' SQL.6: extracted out of former ParseQuery, the SAME shared-helper move
' as ParseOrderByAndLimit just above, for the identical reason.
Private Sub ParseTrailingTerminator(ByVal toks As Collection, ByRef pos As Long)
    If AtOp(toks, pos, ";") Then pos = pos + 1

    Dim leftover As Collection
    Set leftover = CurTok(toks, pos)
    If TokKind(leftover) <> TK_EOF Then
        VLA_Messages.RaiseMsg "sql-unsupported-keyword", "keyword", TokText(leftover)
    End If
End Sub

' Parses queryText against SQL's own frozen grammar (this file's own
' header has it in full). A Sub, not a Boolean-returning Function, since
' failure is always a raised, named refusal here - never a value the
' caller is expected to check, VLA_Datalog.ParseProgram's own shape.
'
' SQL.6: now three calls in sequence (ParseSimpleSelect, then the two
' shared helpers just above) rather than one long inline body - a pure
' internal refactor, byte-for-byte identical parsing behavior to before
' this item (every SQL.1-SQL.5 test needed zero changes to prove it). A
' top-level UNION/INTERSECT/EXCEPT reaching ParseTrailingTerminator here
' still falls through to the ordinary sql-unsupported-keyword refusal,
' UNCHANGED - this function (and SqlRunJoin, its own only caller) stays
' single-query-only forever; SqlRunCompound/ParseCompoundQuery own the
' compound grammar layer instead, never by modifying this one.
Private Sub ParseQuery(ByVal queryText As String, ByRef hasDistinct As Boolean, ByRef selectAll As Boolean, ByRef selectItems As Collection, _
                        ByRef fromName As String, ByRef joins As Collection, ByRef whereExpr As Collection, ByRef hasWhere As Boolean, _
                        ByRef hasGroupBy As Boolean, ByRef groupByItems As Collection, ByRef hasHaving As Boolean, ByRef havingExpr As Collection, _
                        ByRef hasOrderBy As Boolean, ByRef orderItems As Collection, ByRef hasLimit As Boolean, ByRef limitCount As Long)
    Dim toks As Collection
    Set toks = Tokenize(queryText)
    Dim pos As Long
    pos = 1

    ParseSimpleSelect toks, pos, hasDistinct, selectAll, selectItems, fromName, joins, whereExpr, hasWhere, _
                       hasGroupBy, groupByItems, hasHaving, havingExpr
    ParseOrderByAndLimit toks, pos, hasOrderBy, orderItems, hasLimit, limitCount
    ParseTrailingTerminator toks, pos
End Sub

' ---- SQL.6: the compound-query grammar layer - combines multiple
'      simple-SELECT branches with UNION/UNION ALL/INTERSECT/EXCEPT,
'      real SQLite precedence (INTERSECT binds TIGHTER than the other
'      three, which are left-associative with each other), mirroring
'      ParseAddExpr/ParseMulExpr's own precedence-climbing shape - this
'      file's own SQL.6 header note has the full design. -------------

' Leaf: Item(1)=COMPOUND_LEAF, Item(2)=branchText (String, the EXACT
' original-query-text span this one branch's own grammar covered,
' TokStart-sliced - EvalCompoundNode re-runs it through the existing,
' unchanged SqlRunJoin rather than evaluating a parsed AST directly).
Private Function MakeCompoundLeaf(ByVal branchText As String) As Collection
    Dim n As New Collection
    n.Add COMPOUND_LEAF
    n.Add branchText
    Set MakeCompoundLeaf = n
End Function

' Internal node: Item(1)=one of SETOP_UNION/SETOP_UNION_ALL/
' SETOP_INTERSECT/SETOP_EXCEPT, Item(2)=left subtree, Item(3)=right
' subtree - NodeBin's own shape (VLA_Sql.bas's own boolean AST),
' repeated for a completely different tree.
Private Function MakeCompoundSetOp(ByVal opKind As Long, ByVal leftNode As Collection, ByVal rightNode As Collection) As Collection
    Dim n As New Collection
    n.Add opKind
    n.Add leftNode
    n.Add rightNode
    Set MakeCompoundSetOp = n
End Function

Private Function CompoundNodeKind(ByVal n As Collection) As Long
    CompoundNodeKind = n.Item(1)
End Function

Private Function CompoundLeafText(ByVal n As Collection) As String
    CompoundLeafText = n.Item(2)
End Function

Private Function CompoundOpLeft(ByVal n As Collection) As Collection
    Set CompoundOpLeft = n.Item(2)
End Function

Private Function CompoundOpRight(ByVal n As Collection) As Collection
    Set CompoundOpRight = n.Item(3)
End Function

' One simple-SELECT branch: records where it STARTS (the current
' token's own TokStart) before parsing it purely to find where it ENDS
' (pos, after ParseSimpleSelect returns, now sitting at whatever
' terminates this branch's own grammar - a set-op keyword, ORDER BY,
' LIMIT, ';', or EOF) - the parsed AST itself is discarded; only the
' ORIGINAL TEXT SPAN between those two offsets is kept, since evaluation
' re-parses that span from scratch through the existing SqlRunJoin
' (this file's own SQL.6 header note has the "why").
Private Function ParseCompoundLeaf(ByVal toks As Collection, ByRef pos As Long, ByVal queryText As String) As Collection
    Dim branchStart As Long
    branchStart = TokStart(CurTok(toks, pos))

    Dim hasDistinct As Boolean, selectAll As Boolean, selectItems As Collection
    Dim fromName As String, joins As Collection, whereExpr As Collection, hasWhere As Boolean
    Dim hasGroupBy As Boolean, groupByItems As Collection, hasHaving As Boolean, havingExpr As Collection
    ParseSimpleSelect toks, pos, hasDistinct, selectAll, selectItems, fromName, joins, whereExpr, hasWhere, _
                       hasGroupBy, groupByItems, hasHaving, havingExpr

    Dim branchEnd As Long
    branchEnd = TokStart(CurTok(toks, pos))
    Set ParseCompoundLeaf = MakeCompoundLeaf(Mid$(queryText, branchStart, branchEnd - branchStart))
End Function

' intersectChain ::= simpleSelect ('INTERSECT' simpleSelect)* - binds
' TIGHTER than the setOpChain level below it, by construction: that
' level calls this one for each of its own operands, never the reverse
' (ParseMulExpr's own precedence relationship to ParseAddExpr, mirrored).
Private Function ParseIntersectChain(ByVal toks As Collection, ByRef pos As Long, ByVal queryText As String) As Collection
    Dim leftNode As Collection
    Set leftNode = ParseCompoundLeaf(toks, pos, queryText)
    Do While AtKeyword(toks, pos, "intersect")
        pos = pos + 1
        Dim rightNode As Collection
        Set rightNode = ParseCompoundLeaf(toks, pos, queryText)
        Set leftNode = MakeCompoundSetOp(SETOP_INTERSECT, leftNode, rightNode)
    Loop
    Set ParseIntersectChain = leftNode
End Function

' setOpChain ::= intersectChain (('UNION' ['ALL'] | 'EXCEPT') intersectChain)*
' - left-associative, UNION/UNION ALL/EXCEPT all sharing this ONE
' precedence level (real SQLite behavior, verified: `A EXCEPT B UNION C`
' means `(A EXCEPT B) UNION C`, never right-associated or reordered by
' operator identity).
Private Function ParseSetOpChain(ByVal toks As Collection, ByRef pos As Long, ByVal queryText As String) As Collection
    Dim leftNode As Collection
    Set leftNode = ParseIntersectChain(toks, pos, queryText)
    Do While AtKeyword(toks, pos, "union") Or AtKeyword(toks, pos, "except")
        Dim opKind As Long
        If AtKeyword(toks, pos, "union") Then
            pos = pos + 1
            If AtKeyword(toks, pos, "all") Then
                opKind = SETOP_UNION_ALL
                pos = pos + 1
            Else
                opKind = SETOP_UNION
            End If
        Else
            pos = pos + 1
            opKind = SETOP_EXCEPT
        End If
        Dim rightNode As Collection
        Set rightNode = ParseIntersectChain(toks, pos, queryText)
        Set leftNode = MakeCompoundSetOp(opKind, leftNode, rightNode)
    Loop
    Set ParseSetOpChain = leftNode
End Function

' The compound-query entry point's own parse half - tokenizes queryText
' ONCE, builds the whole setOpChain tree, then parses ORDER BY/LIMIT
' (ParseOrderByAndLimit) and the trailing ';'/leftover check
' (ParseTrailingTerminator) exactly once, trailing the WHOLE tree - the
' SAME shared helpers ParseQuery's own single-query path uses, so this
' grammar can never drift from that one. root comes back as EITHER a
' single COMPOUND_LEAF (queryText had no top-level set operator at all -
' SqlRunCompound's own short-circuit to plain SqlRunJoin relies on this)
' or a real SETOP tree.
Private Sub ParseCompoundQuery(ByVal queryText As String, ByRef root As Collection, _
                               ByRef hasOrderBy As Boolean, ByRef orderItems As Collection, _
                               ByRef hasLimit As Boolean, ByRef limitCount As Long)
    Dim toks As Collection
    Set toks = Tokenize(queryText)
    Dim pos As Long
    pos = 1

    Set root = ParseSetOpChain(toks, pos, queryText)
    ParseOrderByAndLimit toks, pos, hasOrderBy, orderItems, hasLimit, limitCount
    ParseTrailingTerminator toks, pos
End Sub

Private Function ParseOrExpr(ByVal toks As Collection, ByRef pos As Long) As Collection
    Dim left As Collection
    Set left = ParseAndExpr(toks, pos)
    Do While AtKeyword(toks, pos, "or")
        pos = pos + 1
        Dim right As Collection
        Set right = ParseAndExpr(toks, pos)
        Set left = NodeBin(NK_OR, left, right)
    Loop
    Set ParseOrExpr = left
End Function

Private Function ParseAndExpr(ByVal toks As Collection, ByRef pos As Long) As Collection
    Dim left As Collection
    Set left = ParseNotExpr(toks, pos)
    Do While AtKeyword(toks, pos, "and")
        pos = pos + 1
        Dim right As Collection
        Set right = ParseNotExpr(toks, pos)
        Set left = NodeBin(NK_AND, left, right)
    Loop
    Set ParseAndExpr = left
End Function

' A real, live-caught grammar bug, not a wiring slip (SQL.1): '(' expr
' ')' must be reachable as a complete boolean expression (NOT (dept =
' 'eng'), (a = 1 OR b = 2) AND c = 3) but must NEVER also be reachable
' as a scalar "primary" (a comparison's own operand) - those are
' different rules that happen to share a bracket character, and
' conflating them here is exactly what broke NOT (dept = 'eng'):
' ParseNotExpr's own non-NOT fallback used to go straight to
' ParseComparison, whose own ParsePrimary swallowed the ENTIRE "(dept =
' 'eng')" as if it were one scalar value, leaving ParseComparison
' hunting for a comparison operator AFTER a group that had already
' completed one internally - "expected a comparison operator..., found
' end of query" on a perfectly well-formed query. SQL.1 had no
' arithmetic, so this used to be simple: EVERY '(' reaching THIS level
' was unambiguously a boolean group. SQL.2 gave ParsePrimary its OWN,
' SEPARATE '(' handling (a parenthesized addExpr) - safe specifically
' because it is reached from DIFFERENT positions than this function
' (deep inside a comparison's operand or a SELECT-list item, never from
' this function's own fallback path, which still tries THIS '(' branch
' first, unconditionally, before ParseComparison ever runs) - this
' function's own claim on a LEADING '(' at the top of a condition is
' UNCHANGED and still wins every time, which is also why a comparison's
' own LEFT operand can never itself start with '(' (VLA_Sql.bas's own
' SQL.2 header note has the full limit).
Private Function ParseNotExpr(ByVal toks As Collection, ByRef pos As Long) As Collection
    If AtKeyword(toks, pos, "not") Then
        pos = pos + 1
        Dim operand As Collection
        Set operand = ParseNotExpr(toks, pos)   ' right-associative - NOT NOT x parses, for free
        Set ParseNotExpr = NodeNot(operand)
    ElseIf AtOp(toks, pos, "(") Then
        pos = pos + 1
        Dim inner As Collection
        Set inner = ParseOrExpr(toks, pos)
        ExpectOp toks, pos, ")"
        Set ParseNotExpr = inner
    Else
        Set ParseNotExpr = ParseComparison(toks, pos)
    End If
End Function

' Every WHERE leaf in SQL's own frozen subset is a comparison - there
' is no bare "truthy column" test (unlike some SQL dialects' own boolean
' columns), so an addExpr with no following comparison operator is
' refused here, by name, rather than silently treated as always-true.
' SQL.2: operands are addExpr, not a bare primary - WHERE salary * 1.1
' > 100000 now parses with no parens needed, the same evaluator (this
' file's own header note) reused rather than growing a second scalar
' grammar just for the projection list.
Private Function ParseComparison(ByVal toks As Collection, ByRef pos As Long) As Collection
    Dim left As Collection
    Set left = ParseAddExpr(toks, pos)
    Dim t As Collection
    Set t = CurTok(toks, pos)
    If TokKind(t) = TK_OP Then
        Select Case TokText(t)
        Case "=", "<>", "<", "<=", ">", ">="
            Dim op As String
            op = TokText(t)
            pos = pos + 1
            Dim right As Collection
            Set right = ParseAddExpr(toks, pos)
            Set ParseComparison = NodeCmp(op, left, right)
            Exit Function
        End Select
    End If
    VLA_Messages.RaiseMsg "sql-expected-token", "expected", "a comparison operator (=, <>, <, <=, >, >=)", "found", DescribeTok(CurTok(toks, pos))
End Function

' SQL.2: addExpr ::= mulExpr (('+' | '-') mulExpr)* - left-associative,
' the standard arithmetic precedence level ABOVE multiplication/
' division. Two callers: a comparison's own operand (ParseComparison,
' above) and a SELECT-list projection item (ParseQuery, below) - one
' evaluator, pointed at two places, this file's own header note made
' real for arithmetic the same way it already was for comparisons.
Private Function ParseAddExpr(ByVal toks As Collection, ByRef pos As Long) As Collection
    Dim left As Collection
    Set left = ParseMulExpr(toks, pos)
    Do While AtOp(toks, pos, "+") Or AtOp(toks, pos, "-")
        Dim op As String
        op = TokText(CurTok(toks, pos))
        pos = pos + 1
        Dim right As Collection
        Set right = ParseMulExpr(toks, pos)
        Set left = NodeArith(op, left, right)
    Loop
    Set ParseAddExpr = left
End Function

' mulExpr ::= primary (('*' | '/') primary)* - binds TIGHTER than
' addExpr (a * b + c means (a * b) + c), by construction: addExpr calls
' this for each of its own operands, never the reverse.
Private Function ParseMulExpr(ByVal toks As Collection, ByRef pos As Long) As Collection
    Dim left As Collection
    Set left = ParsePrimary(toks, pos)
    Do While AtOp(toks, pos, "*") Or AtOp(toks, pos, "/")
        Dim op As String
        op = TokText(CurTok(toks, pos))
        pos = pos + 1
        Dim right As Collection
        Set right = ParsePrimary(toks, pos)
        Set left = NodeArith(op, left, right)
    Loop
    Set ParseMulExpr = left
End Function

' Scalar operands - a column, a number, a string, a parenthesized
' addExpr, or a unary +/-. SQL.2 added '(' handling HERE, deliberately
' never at ParseNotExpr's own level (this file's own SQL.2 header note
' has the full reasoning: these are reached from DIFFERENT grammar
' positions - deep inside a comparison's operand or a SELECT-list item,
' never from cond's own top level - so there is no repeat of the
' original NOT (x = 'y') conflation; ParseNotExpr's own '(' check,
' unchanged, always wins the race for a leading '(' at the very top of
' a condition, before ParseComparison/ParseAddExpr/ParsePrimary ever
' get a chance to run there at all). Unary minus is (0 - operand),
' reusing NK_ARITH's own binary "-" evaluator path rather than a
' unary-specific AST node/evaluator case; unary plus is a pure no-op,
' added for symmetry at zero extra evaluator cost.
Private Function ParsePrimary(ByVal toks As Collection, ByRef pos As Long) As Collection
    If AtOp(toks, pos, "(") Then
        pos = pos + 1
        Dim inner As Collection
        Set inner = ParseAddExpr(toks, pos)
        ExpectOp toks, pos, ")"
        Set ParsePrimary = inner
        Exit Function
    ElseIf AtOp(toks, pos, "-") Then
        pos = pos + 1
        Dim negOperand As Collection
        Set negOperand = ParsePrimary(toks, pos)
        Set ParsePrimary = NodeArith("-", NodeNumber(0#), negOperand)
        Exit Function
    ElseIf AtOp(toks, pos, "+") Then
        pos = pos + 1
        Set ParsePrimary = ParsePrimary(toks, pos)
        Exit Function
    End If
    Dim t As Collection
    Set t = CurTok(toks, pos)
    Select Case TokKind(t)
    Case TK_IDENT
        ' SQL.4: an aggregate call - COUNT/SUM/MIN/MAX/AVG followed
        ' IMMEDIATELY by '(' - keyword-in-context, checked BEFORE the
        ' ordinary column/qualified-column path below, so a column
        ' genuinely named "count" (never followed by '(') is completely
        ' unaffected. This file's own SQL.4 header note has the full
        ' grammar/AST reasoning.
        Dim aggKind As Long
        aggKind = AggKindForKeyword(VLA_Identity.Fold(TokText(t)))
        If aggKind <> -1 And AtOp(toks, pos + 1, "(") Then
            pos = pos + 2   ' consume the function name and '('
            Dim aggIsStar As Boolean, aggOperand As Collection
            aggIsStar = False
            Set aggOperand = Nothing
            If AtOp(toks, pos, "*") Then
                If aggKind <> VLA_Relation.AGG_COUNT Then
                    VLA_Messages.RaiseMsg "sql-aggregate-star-only-count", "function", UCase$(TokText(t))
                End If
                aggIsStar = True
                pos = pos + 1
            Else
                Set aggOperand = ParseAddExpr(toks, pos)
                If NodeContainsAggregate(aggOperand) Then
                    VLA_Messages.RaiseMsg "sql-aggregate-nested"
                End If
            End If
            ExpectOp toks, pos, ")"
            Set ParsePrimary = NodeAgg(aggKind, aggIsStar, aggOperand)
            Exit Function
        End If
        pos = pos + 1
        ' SQL.3: a qualified column reference, Employees.id - the '.'
        ' tokenizes as its own TK_OP (Tokenize's own header note), so an
        ' identifier immediately followed by one is reassembled here,
        ' consuming three tokens instead of one, into a single NK_COLUMN
        ' node carrying both parts. Anything else after the identifier
        ' (no '.', or a '.' not followed by a second identifier) leaves
        ' it an ordinary unqualified column - unchanged from SQL.1/SQL.2.
        If AtOp(toks, pos, ".") Then
            Dim afterDot As Collection
            Set afterDot = CurTok(toks, pos + 1)
            If TokKind(afterDot) = TK_IDENT Then
                pos = pos + 2
                Set ParsePrimary = NodeColumn(VLA_Identity.Fold(TokText(afterDot)), TokText(afterDot), VLA_Identity.Fold(TokText(t)))
                Exit Function
            End If
        End If
        Set ParsePrimary = NodeColumn(VLA_Identity.Fold(TokText(t)), TokText(t))
    Case TK_NUMBER
        pos = pos + 1
        Set ParsePrimary = NodeNumber(ParseInvariantNumber(TokText(t)))
    Case TK_STRING
        pos = pos + 1
        Set ParsePrimary = NodeString(TokText(t))
    Case Else
        VLA_Messages.RaiseMsg "sql-expected-token", "expected", "a column, number, string, or '('", "found", DescribeTok(t)
    End Select
End Function

' ---- SQL.4: aggregate-call grammar/AST helpers -------------------------

' kw must already be folded (VLA_Identity.Fold) - returns -1 for
' anything that isn't one of the five aggregate function names, so
' ParsePrimary's own caller can tell "not an aggregate call" apart from
' a real aggregate kind (0 is a valid kind, AGG_COUNT).
Private Function AggKindForKeyword(ByVal kw As String) As Long
    Select Case kw
    Case "count": AggKindForKeyword = VLA_Relation.AGG_COUNT
    Case "sum": AggKindForKeyword = VLA_Relation.AGG_SUM
    Case "min": AggKindForKeyword = VLA_Relation.AGG_MIN
    Case "max": AggKindForKeyword = VLA_Relation.AGG_MAX
    Case "avg": AggKindForKeyword = VLA_Relation.AGG_AVG
    Case Else: AggKindForKeyword = -1
    End Select
End Function

Private Function AggFunctionName(ByVal kind As Long) As String
    Select Case kind
    Case VLA_Relation.AGG_COUNT: AggFunctionName = "COUNT"
    Case VLA_Relation.AGG_SUM: AggFunctionName = "SUM"
    Case VLA_Relation.AGG_MIN: AggFunctionName = "MIN"
    Case VLA_Relation.AGG_MAX: AggFunctionName = "MAX"
    Case VLA_Relation.AGG_AVG: AggFunctionName = "AVG"
    End Select
End Function

' Whether node's own tree contains an aggregate call ANYWHERE - walks
' both scalar node kinds (NK_ARITH/NK_AGG) and boolean node kinds
' (NK_CMP/NK_AND/NK_OR/NK_NOT), since this is used against WHERE/ON
' (boolean trees) as well as an aggregate's own operand and a SELECT
' item (scalar trees). Not short-circuit (Or, like EvalBool's own AND/OR
' cases) - harmless, no side effects anywhere in this subset.
Private Function NodeContainsAggregate(ByVal node As Collection) As Boolean
    Select Case NodeKind(node)
    Case NK_AGG
        NodeContainsAggregate = True
    Case NK_ARITH
        NodeContainsAggregate = NodeContainsAggregate(NodeArithLeft(node)) Or NodeContainsAggregate(NodeArithRight(node))
    Case NK_CMP
        NodeContainsAggregate = NodeContainsAggregate(NodeCmpLeft(node)) Or NodeContainsAggregate(NodeCmpRight(node))
    Case NK_AND, NK_OR
        NodeContainsAggregate = NodeContainsAggregate(NodeBinLeft(node)) Or NodeContainsAggregate(NodeBinRight(node))
    Case NK_NOT
        NodeContainsAggregate = NodeContainsAggregate(NodeNotOperand(node))
    End Select
End Function

' A canonical, internal-only text form of a scalar expression - never
' shown to the user, only ever used as a Dictionary key (AggSignature's
' own dedup, and the grouped colMap's own aggregate-column key, below),
' so CStr's own locale-dependent number formatting is harmless here: both
' the key's registration and every lookup against it happen within the
' SAME call, same locale, same process.
Private Function ExprSignature(ByVal node As Collection) As String
    Select Case NodeKind(node)
    Case NK_COLUMN
        If Len(NodeColumnQualifier(node)) > 0 Then
            ExprSignature = NodeColumnQualifier(node) & "." & NodeColumnName(node)
        Else
            ExprSignature = NodeColumnName(node)
        End If
    Case NK_NUMBER
        ExprSignature = "#" & CStr(NodeNumberValue(node))
    Case NK_STRING
        ExprSignature = "'" & NodeStringValue(node) & "'"
    Case NK_ARITH
        ExprSignature = "(" & ExprSignature(NodeArithLeft(node)) & NodeArithOp(node) & ExprSignature(NodeArithRight(node)) & ")"
    Case NK_AGG
        ExprSignature = AggSignature(node)
    End Select
End Function

' An aggregate call's own canonical signature (e.g. "count(*)",
' "sum(salary)") - doubles as BOTH the dedup key ComputeGroupedRows uses
' to collapse the same aggregate named twice (once in SELECT, once in
' HAVING) into one computed slot, AND the grouped colMap's own key for
' that aggregate's result column - EvalScalar's own new NK_AGG case
' (below) looks a node back up by this SAME signature, so registration
' and lookup can never drift apart.
Private Function AggSignature(ByVal node As Collection) As String
    If NodeAggIsStar(node) Then
        AggSignature = LCase$(AggFunctionName(NodeAggKind(node))) & "(*)"
    Else
        AggSignature = LCase$(AggFunctionName(NodeAggKind(node))) & "(" & ExprSignature(NodeAggOperand(node)) & ")"
    End If
End Function

' A default output header for an un-aliased aggregate call (e.g.
' "COUNT(*)", "SUM(salary)") - this file's own SQL.4 header note has the
' reasoning for why an aggregate call, unlike an arithmetic expression,
' gets one at all.
Private Function AggDisplayText(ByVal node As Collection) As String
    If NodeAggIsStar(node) Then
        AggDisplayText = AggFunctionName(NodeAggKind(node)) & "(*)"
    Else
        AggDisplayText = AggFunctionName(NodeAggKind(node)) & "(" & ExprSignature(NodeAggOperand(node)) & ")"
    End If
End Function

' Exact structural equality between two column references (both already
' folded at parse time) - the ONLY shape a GROUP BY item can be (this
' file's own SQL.4 header note), so this is also the ONLY comparison
' SQL.4's own "every SELECT item must be a GROUP BY column or an
' aggregate" rule needs. Returns False (never raises/crashes) whenever
' either node isn't a column at all - a non-column SELECT item (an
' arithmetic expression, say) can never match a GROUP BY item this way,
' which is exactly the intended refusal.
Private Function ColumnNodesEqual(ByVal a As Collection, ByVal b As Collection) As Boolean
    If NodeKind(a) <> NK_COLUMN Or NodeKind(b) <> NK_COLUMN Then Exit Function
    ColumnNodesEqual = (StrComp(NodeColumnName(a), NodeColumnName(b), vbBinaryCompare) = 0) _
                    And (StrComp(NodeColumnQualifier(a), NodeColumnQualifier(b), vbBinaryCompare) = 0)
End Function

' Whether ANY SELECT-list item contains an aggregate call anywhere in
' its own tree - part of "does this query aggregate at all" (SqlRunJoin
' and ParseQuery both ask this, e.g. SELECT COUNT(*) FROM t with no
' GROUP BY clause at all still aggregates, real SQL's own rule).
Private Function AnySelectItemHasAggregate(ByVal selectItems As Collection) As Boolean
    Dim i As Long
    For i = 1 To selectItems.Count
        If NodeContainsAggregate(ProjItemExpr(selectItems.Item(i))) Then
            AnySelectItemHasAggregate = True
            Exit Function
        End If
    Next i
End Function

' Walks node (a scalar OR boolean tree - reused for both a SELECT item
' and HAVING's own boolean expression), collecting every DISTINCT
' aggregate call (AggSignature's own dedup, via seen) into outAgg, in
' first-seen order - the ordering ComputeGroupedRows' own aggSpecs and
' the grouped colMap's own result-column positions both rely on.
Private Sub CollectAggregateNodes(ByVal node As Collection, ByVal seen As Object, ByVal outAgg As Collection)
    Select Case NodeKind(node)
    Case NK_AGG
        Dim sig As String
        sig = AggSignature(node)
        If Not VLA_Runtime.VlaDictHas(seen, sig) Then
            VLA_Runtime.VlaDictSet seen, sig, True
            outAgg.Add node
        End If
    Case NK_ARITH
        CollectAggregateNodes NodeArithLeft(node), seen, outAgg
        CollectAggregateNodes NodeArithRight(node), seen, outAgg
    Case NK_CMP
        CollectAggregateNodes NodeCmpLeft(node), seen, outAgg
        CollectAggregateNodes NodeCmpRight(node), seen, outAgg
    Case NK_AND, NK_OR
        CollectAggregateNodes NodeBinLeft(node), seen, outAgg
        CollectAggregateNodes NodeBinRight(node), seen, outAgg
    Case NK_NOT
        CollectAggregateNodes NodeNotOperand(node), seen, outAgg
    End Select
End Sub

' ---- scalar/boolean evaluator - SQL.2's computed columns, SQL.3's
'      residual ON conditions, and SQL.4's HAVING all point this SAME
'      evaluator at different rows, rather than growing a second one ---
'
' IsNumericValue/CompareValues used to live here as Private functions;
' DATALOG.4 hoisted both into VLA_Relation.bas (ValueIsNumericType/
' CompareValues) the moment a second engine needed the identical
' mechanical compare - TableArgResolve's own SQL.1 precedent, repeated.
' EvalBool's own NK_CMP case below now calls the shared function
' directly, passing ValueIsNumericType(l) And ValueIsNumericType(r) as
' its own bothNumeric verdict - SQL.1's STRICT policy, unchanged,
' byte-for-byte: the shared function's own numeric branch only ever
' runs here when BOTH sides are already a real Excel Value2 numeric
' type, exactly as before this hoist.
' SQL.3's own single chokepoint for turning a (possibly-qualified)
' NK_COLUMN reference into an absolute position in the CURRENT joined
' row - colMap (SqlRunJoin's own AddColumnsToMap builds it) holds BOTH
' kinds of key for every column registered so far: "tablename.colname"
' (always precise) and, when that folded column name is still unique
' across every table registered so far, the bare "colname" too. A bare
' name that exists in MORE than one joined table is never silently
' pointed at whichever table happened to register it first (or last) -
' AddColumnsToMap marks it Empty instead of overwriting the earlier
' position, and THIS function is what turns that sentinel into a named
' refusal (sql-ambiguous-column) rather than a wrong answer. A single-
' table query (SQL.1/SQL.2, or a SQL.3 query with no JOIN at all) never
' hits either branch here - one table can never collide with itself, so
' the bare "colname" key always resolves, unqualified, exactly as before.
Private Function ResolveColumnPos(ByVal node As Collection, ByVal colMap As Object) As Long
    Dim qual As String
    qual = NodeColumnQualifier(node)
    Dim nm As String
    nm = NodeColumnName(node)
    If Len(qual) > 0 Then
        Dim qkey As String
        qkey = qual & "." & nm
        If Not VLA_Runtime.VlaDictHas(colMap, qkey) Then VLA_Messages.RaiseMsg "sql-unknown-column", "column", qual & "." & nm
        ResolveColumnPos = CLng(VLA_Runtime.VlaDictGet(colMap, qkey))
    Else
        If Not VLA_Runtime.VlaDictHas(colMap, nm) Then VLA_Messages.RaiseMsg "sql-unknown-column", "column", nm
        Dim v As Variant
        v = VLA_Runtime.VlaDictGet(colMap, nm)
        If IsEmpty(v) Then VLA_Messages.RaiseMsg "sql-ambiguous-column", "column", nm
        ResolveColumnPos = CLng(v)
    End If
End Function

' SQL.4's own chokepoint for turning an aggregate call back into its
' own absolute position in the GROUPED row - colMap here is
' ComputeGroupedRows' own groupColMap, keyed by AggSignature for every
' aggregate this query actually needs (registered before any row is
' ever evaluated). Never a user-facing refusal - a lookup miss here
' means ComputeGroupedRows' own enumeration and this node disagree,
' which can only be an internal bug, not a bad query (EvalScalar's own
' "Case Else" below uses the identical Err.Raise 5 convention for the
' same reason).
Private Function ResolveAggPos(ByVal node As Collection, ByVal colMap As Object) As Long
    Dim key As String
    key = AggSignature(node)
    If Not VLA_Runtime.VlaDictHas(colMap, key) Then
        Err.Raise 5, "VLA-Sql", "internal: aggregate '" & key & "' not registered in the grouped colMap"
    End If
    ResolveAggPos = CLng(VLA_Runtime.VlaDictGet(colMap, key))
End Function

Private Function EvalScalar(ByVal node As Collection, ByVal colMap As Object, ByRef row() As Variant) As Variant
    Select Case NodeKind(node)
    Case NK_COLUMN
        EvalScalar = row(ResolveColumnPos(node, colMap))
    Case NK_NUMBER
        EvalScalar = NodeNumberValue(node)
    Case NK_STRING
        EvalScalar = NodeStringValue(node)
    Case NK_ARITH
        ' SQL.2: ComputeArithmetic is DATALOG.4's own hoist into
        ' VLA_Relation.bas, getting its SECOND real caller here - SQL's
        ' own numeric-ness policy stays STRICT (ValueIsNumericType
        ' alone), the identical policy EvalBool's own NK_CMP case
        ' already uses below, never DATALOG's own more lenient
        ' numeric-looking-string policy (SQL's values are always real
        ' Excel Value2 reads, never a hand-written rule-text constant).
        Dim al As Variant, ar As Variant
        al = EvalScalar(NodeArithLeft(node), colMap, row)
        ar = EvalScalar(NodeArithRight(node), colMap, row)
        Dim okArith As Boolean, reasonArith As String
        EvalScalar = VLA_Relation.ComputeArithmetic(NodeArithOp(node), al, ar, _
            VLA_Relation.ValueIsNumericType(al) And VLA_Relation.ValueIsNumericType(ar), okArith, reasonArith)
        If Not okArith Then
            If reasonArith = "divide-by-zero" Then
                VLA_Messages.RaiseMsg "sql-division-by-zero"
            Else
                VLA_Messages.RaiseMsg "sql-arithmetic-non-numeric-operand", "operator", NodeArithOp(node)
            End If
        End If
    Case NK_AGG
        ' SQL.4: an aggregate call is a scalar VALUE, exactly like a
        ' column or a number - reached only when colMap/row are
        ' ComputeGroupedRows' own grouped colMap/row (SELECT-list items
        ' and HAVING, post-grouping); the operand itself is NEVER
        ' re-evaluated here (it was already folded into the grouped
        ' row's own value by RelGroupBy, against the SOURCE row, before
        ' this ever runs).
        EvalScalar = row(ResolveAggPos(node, colMap))
    Case Else
        Err.Raise 5, "VLA-Sql", "internal: EvalScalar called on a non-scalar AST node"
    End Select
End Function

' SQL.2: walks a scalar expression tree checking every NK_COLUMN
' reference resolves against colMap - called UPFRONT, once per SELECT-
' list item (SqlRun, below), BEFORE the row loop, so an unknown column
' in the SELECT list is refused REGARDLESS of how many rows the table
' actually has (SQL.1's own existing guarantee for the SELECT list,
' tested by TestSql's own "unknown SELECT column" case - preserved
' here, not just re-derived by accident, since EvalScalar's OWN
' NK_COLUMN check only ever runs DURING evaluation, which a zero-row
' table would never reach).
Private Sub ValidateScalarColumns(ByVal node As Collection, ByVal colMap As Object)
    Select Case NodeKind(node)
    Case NK_COLUMN
        Call ResolveColumnPos(node, colMap)
    Case NK_ARITH
        ValidateScalarColumns NodeArithLeft(node), colMap
        ValidateScalarColumns NodeArithRight(node), colMap
    Case NK_AGG
        ' SQL.4: validates that THIS node resolves against colMap - the
        ' aggregate's own OPERAND is deliberately NOT walked here (unlike
        ' NK_ARITH's own recursion into its two operands): reached only
        ' against the GROUPED colMap (post-grouping SELECT-list items/
        ' HAVING), where the operand's own source columns are no longer
        ' meaningful at all - the operand was already validated
        ' separately, against the SOURCE colMap, by ComputeGroupedRows
        ' itself before any grouping ever ran.
        Call ResolveAggPos(node, colMap)
    End Select
End Sub

' Not short-circuit (VBA's And/Or never are) - harmless here, since
' SQL.1's own frozen subset has no side effects and no operation that
' can raise on a value it would otherwise skip (unlike, say, division -
' not part of this subset either), so evaluating both sides of AND/OR
' unconditionally always produces the correct boolean, just occasionally
' more work than strictly needed.
Private Function EvalBool(ByVal node As Collection, ByVal colMap As Object, ByRef row() As Variant) As Boolean
    Select Case NodeKind(node)
    Case NK_CMP
        Dim l As Variant, r As Variant
        l = EvalScalar(NodeCmpLeft(node), colMap, row)
        r = EvalScalar(NodeCmpRight(node), colMap, row)
        EvalBool = VLA_Relation.CompareValues(NodeCmpOp(node), l, r, VLA_Relation.ValueIsNumericType(l) And VLA_Relation.ValueIsNumericType(r))
    Case NK_AND
        EvalBool = EvalBool(NodeBinLeft(node), colMap, row) And EvalBool(NodeBinRight(node), colMap, row)
    Case NK_OR
        EvalBool = EvalBool(NodeBinLeft(node), colMap, row) Or EvalBool(NodeBinRight(node), colMap, row)
    Case NK_NOT
        EvalBool = Not EvalBool(NodeNotOperand(node), colMap, row)
    Case Else
        Err.Raise 5, "VLA-Sql", "internal: EvalBool called on a non-boolean AST node"
    End Select
End Function

' A manually-doubled growable buffer - BETA_ROADMAP2.md's own SQL.1
' words, named because per-row ReDim Preserve is a genuine O(n^2) trap a
' large table (this engine's whole stated motivation) would actually
' hit, not a hypothetical one. cap/used are tracked by the caller
' alongside buf() rather than probed via UBound each call, since UBound
' on a never-yet-ReDim'd array raises - avoided entirely this way rather
' than caught.
Private Sub BufAdd(ByRef buf() As Variant, ByRef used As Long, ByRef cap As Long, ByVal item As Variant)
    If used >= cap Then
        Dim newCap As Long
        If cap = 0 Then
            newCap = 8
            ReDim buf(1 To newCap)
        Else
            newCap = cap * 2
            ReDim Preserve buf(1 To newCap)
        End If
        cap = newCap
    End If
    used = used + 1
    buf(used) = item
End Sub

Private Function ColNameFolded(ByVal pair As Collection) As String
    ColNameFolded = pair.Item(1)
End Function

Private Function ColNameOriginal(ByVal pair As Collection) As String
    ColNameOriginal = pair.Item(2)
End Function

' ---- table records (SQL.3): Item(1)=name (String, folded - the same
'      real-table-name identity FROM/JOIN clauses resolve against),
'      Item(2)=columnNames (Collection, VLA_Relation.RangeColumnNames'
'      own (folded, original) pair shape), Item(3)=rows (Collection,
'      VLA_Relation.RangeToRows' own bag-of-tuple-arrays shape - never
'      a Relation, SQL.1's own bag/set distinction unchanged) ---------
Private Function MakeTableRec(ByVal name As String, ByVal columnNames As Collection, ByVal rows As Collection) As Collection
    Dim r As New Collection
    r.Add name
    r.Add columnNames
    r.Add rows
    Set MakeTableRec = r
End Function

Private Function TableRecName(ByVal r As Collection) As String
    TableRecName = r.Item(1)
End Function

Private Function TableRecColumns(ByVal r As Collection) As Collection
    Set TableRecColumns = r.Item(2)
End Function

Private Function TableRecRows(ByVal r As Collection) As Collection
    Set TableRecRows = r.Item(3)
End Function

' SQL.3: registers one table's own columns into colMap at ABSOLUTE
' positions [baseOffset+1 .. baseOffset+arity] in the eventual fully-
' joined row - both a QUALIFIED key ("tablename.colname", always
' precise) and, when the folded column name is still unique across
' every table registered so far, an UNQUALIFIED key too. A name that
' COLLIDES with an already-registered unqualified entry is marked Empty
' (ResolveColumnPos's own ambiguity sentinel, sql-ambiguous-column)
' rather than silently overwritten - real SQL's own "qualify it"
' requirement, not a guess either way. Also appends this table's own
' ORIGINAL column names, in order, to outHeadersOriginal - SELECT *'s
' own header source once every table has folded in. A single-table
' query never exercises the collision branch at all (one table can
' never collide with itself), so this degenerates to exactly SQL.1/
' SQL.2's own flat colMap in that case.
Private Sub AddColumnsToMap(ByVal colMap As Object, ByVal tableFolded As String, ByVal columnNames As Collection, _
                             ByVal baseOffset As Long, ByVal outHeadersOriginal As Collection)
    Dim i As Long
    For i = 1 To columnNames.Count
        Dim pair As Collection
        Set pair = columnNames.Item(i)
        Dim colFolded As String
        colFolded = ColNameFolded(pair)
        Dim absPos As Long
        absPos = baseOffset + i
        VLA_Runtime.VlaDictSet colMap, tableFolded & "." & colFolded, absPos
        If VLA_Runtime.VlaDictHas(colMap, colFolded) Then
            VLA_Runtime.VlaDictSet colMap, colFolded, Empty
        Else
            VLA_Runtime.VlaDictSet colMap, colFolded, absPos
        End If
        outHeadersOriginal.Add ColNameOriginal(pair)
    Next i
End Sub

' SQL.3's own equi-condition/residual split, the roadmap's own words:
' "the parser splits each ON into equi-conditions (fed to the hash join
' as keys) plus residual non-equi conditions (evaluated as a post-join
' filter through SQL.1's evaluator - zero new evaluation code)." Walks
' onNode (an already-parsed ParseOrExpr AST) through NK_AND at the top
' only - an OR, a NOT, or a comparison whose own operand isn't a bare
' column (an arithmetic expression, a literal, a column from the SAME
' side twice) can never become a hash-join key between these two SPECIFIC
' relations, so the WHOLE such node is kept as one undecomposed residual
' term instead (still fully evaluated via EvalBool, just not accelerated
' as a join key - correctness never depends on how many terms became
' keys; RelJoin's own header note already covers the zero-equi-condition
' case, degenerating correctly to a full cross product before the
' residual filter runs). equiLeft()/equiRight() come back as LOCAL
' 1-based positions within EACH side (RelJoin's own contract) - an
' absolute combined-row position P resolved via colMap is on the
' already-accumulated side, unchanged, when P <= accumArity, or on the
' newly-joined table's own side, at local position P - accumArity,
' otherwise.
Private Sub SplitOnExpr(ByVal onNode As Collection, ByVal colMap As Object, ByVal accumArity As Long, _
                         ByRef equiLeft As Collection, ByRef equiRight As Collection, ByRef residualNode As Collection)
    Set equiLeft = New Collection
    Set equiRight = New Collection
    Dim residualTerms As New Collection
    WalkOnTerm onNode, colMap, accumArity, equiLeft, equiRight, residualTerms
    Set residualNode = Nothing
    Dim k As Long
    For k = 1 To residualTerms.Count
        If residualNode Is Nothing Then
            Set residualNode = residualTerms.Item(k)
        Else
            Set residualNode = NodeBin(NK_AND, residualNode, residualTerms.Item(k))
        End If
    Next k
End Sub

Private Sub WalkOnTerm(ByVal node As Collection, ByVal colMap As Object, ByVal accumArity As Long, _
                        ByVal equiLeft As Collection, ByVal equiRight As Collection, ByVal residualTerms As Collection)
    If NodeKind(node) = NK_AND Then
        WalkOnTerm NodeBinLeft(node), colMap, accumArity, equiLeft, equiRight, residualTerms
        WalkOnTerm NodeBinRight(node), colMap, accumArity, equiLeft, equiRight, residualTerms
        Exit Sub
    End If
    If NodeKind(node) = NK_CMP Then
        If NodeCmpOp(node) = "=" Then
            Dim l As Collection, r As Collection
            Set l = NodeCmpLeft(node)
            Set r = NodeCmpRight(node)
            If NodeKind(l) = NK_COLUMN And NodeKind(r) = NK_COLUMN Then
                Dim lp As Long, rp As Long
                lp = ResolveColumnPos(l, colMap)
                rp = ResolveColumnPos(r, colMap)
                Dim lOnAccum As Boolean, rOnAccum As Boolean
                lOnAccum = (lp <= accumArity)
                rOnAccum = (rp <= accumArity)
                If lOnAccum And Not rOnAccum Then
                    equiLeft.Add lp
                    equiRight.Add rp - accumArity
                    Exit Sub
                ElseIf rOnAccum And Not lOnAccum Then
                    equiLeft.Add rp
                    equiRight.Add lp - accumArity
                    Exit Sub
                End If
            End If
        End If
    End If
    residualTerms.Add node
End Sub

' SQL.4's own bridge into VLA_Relation.RelGroupBy - this file's own
' SQL.4 header note has the full design; VLA_Relation.bas's own header
' has the kernel's own design. colMap/accumRows are the JOIN-folded,
' pre-grouping ones (SqlRunJoin's own colMap/accumRows at the moment
' this is called - WHERE has NOT been applied yet, this function applies
' it itself, once, while building each "eval row"). outColMap comes back
' keyed exactly the way AddColumnsToMap already keys a real table's own
' columns for every GROUP BY column (qualified always; unqualified when
' not already claimed, Empty on collision - ResolveColumnPos's own
' existing ambiguity check handles that sentinel with zero new code),
' plus each aggregate's own canonical AggSignature text for its own
' result column.
Private Function ComputeGroupedRows(ByVal selectItems As Collection, ByVal groupByItems As Collection, _
                                     ByVal hasGroupBy As Boolean, ByVal hasHaving As Boolean, ByVal havingExpr As Collection, _
                                     ByVal hasWhere As Boolean, ByVal whereExpr As Collection, _
                                     ByVal colMap As Object, ByVal accumRows As Collection, _
                                     ByRef outColMap As Object) As Collection
    Dim groupArity As Long
    groupArity = groupByItems.Count

    Dim gi As Long
    For gi = 1 To groupArity
        Call ResolveColumnPos(groupByItems.Item(gi), colMap)
    Next gi

    ' Enumerate every DISTINCT aggregate call across the SELECT list and
    ' HAVING, first-seen order - CollectAggregateNodes' own AggSignature
    ' dedup, so the same call named twice (once in SELECT, once in
    ' HAVING) shares one computed slot rather than two.
    Dim aggList As New Collection
    Dim aggSeen As Object
    Set aggSeen = VLA_Runtime.VlaDictNew()
    Dim si As Long
    For si = 1 To selectItems.Count
        CollectAggregateNodes ProjItemExpr(selectItems.Item(si)), aggSeen, aggList
    Next si
    If hasHaving Then CollectAggregateNodes havingExpr, aggSeen, aggList

    ' Each non-COUNT aggregate gets its own "eval row" value-slot,
    ' positions groupArity+1.. - COUNT needs none at all (VLA_Relation.
    ' bas' own header note has the reason: COUNT(*) and COUNT(col)
    ' degenerate to the identical group row-count here). Every
    ' aggregate's own operand is still validated upfront against colMap
    ' regardless (an unknown column inside COUNT(nosuchcol) is refused
    ' the same way an unknown SELECT column already is), even when its
    ' own value is never evaluated per row.
    Dim aggSpecs As New Collection
    Dim slotOf() As Long
    If aggList.Count > 0 Then ReDim slotOf(1 To aggList.Count)
    Dim nextSlot As Long
    nextSlot = groupArity
    Dim ak As Long
    Dim aNode As Collection
    For ak = 1 To aggList.Count
        Set aNode = aggList.Item(ak)
        If NodeAggKind(aNode) = VLA_Relation.AGG_COUNT Then
            slotOf(ak) = 0
            If Not NodeAggIsStar(aNode) Then ValidateScalarColumns NodeAggOperand(aNode), colMap
            aggSpecs.Add VLA_Relation.MakeAggSpec(VLA_Relation.AGG_COUNT, 0)
        Else
            ValidateScalarColumns NodeAggOperand(aNode), colMap
            nextSlot = nextSlot + 1
            slotOf(ak) = nextSlot
            aggSpecs.Add VLA_Relation.MakeAggSpec(NodeAggKind(aNode), nextSlot)
        End If
    Next ak
    Dim evalArity As Long
    evalArity = nextSlot

    ' WHERE filters BEFORE grouping (real SQL's own order, one step
    ' earlier than DISTINCT's existing "WHERE filters before DISTINCT,
    ' not after" - SQL.2's own precedent, repeated); each surviving row
    ' becomes one "eval row" - its own GROUP BY key values, then each
    ' NEEDED aggregate operand's own value, both evaluated against the
    ' SOURCE colMap/row, exactly like an ordinary WHERE/SELECT column
    ' reference always has been.
    Dim evalRows As New Collection
    Dim t As Variant, arr() As Variant
    Dim keep As Boolean
    Dim evRow() As Variant
    Dim i As Long
    For Each t In accumRows
        arr = t
        keep = True
        If hasWhere Then keep = EvalBool(whereExpr, colMap, arr)
        If keep Then
            ' evalArity can be 0 (e.g. a bare SELECT COUNT(*) FROM t,
            ' no GROUP BY, no non-COUNT aggregate) - a direct
            ' ReDim evRow(1 To 0) is the exact inverted-bounds trap
            ' RelUnit's own header (VLA_Relation.bas) already documents
            ' as unreliable on real Windows Excel (live-caught runtime
            ' error 9); Array() is the reliable empty-array construction
            ' there, reused here the same way.
            If evalArity = 0 Then
                evRow = Array()
            Else
                ReDim evRow(1 To evalArity)
            End If
            For i = 1 To groupArity
                evRow(i) = EvalScalar(groupByItems.Item(i), colMap, arr)
            Next i
            For ak = 1 To aggList.Count
                If slotOf(ak) > 0 Then
                    evRow(slotOf(ak)) = EvalScalar(NodeAggOperand(aggList.Item(ak)), colMap, arr)
                End If
            Next ak
            evalRows.Add evRow
        End If
    Next t

    Dim keyPosColl As New Collection
    For gi = 1 To groupArity
        keyPosColl.Add gi
    Next gi
    Dim keyPositions As Variant
    keyPositions = VLA_Relation.CollToLongArray(keyPosColl)

    Dim gOk As Boolean, gReason As String, gFailIdx As Long
    Dim groupedRows As Collection
    Set groupedRows = VLA_Relation.RelGroupBy(evalRows, keyPositions, aggSpecs, Not hasGroupBy, gOk, gReason, gFailIdx)
    If Not gOk Then
        Dim failNode As Collection
        Set failNode = aggList.Item(gFailIdx)
        If gReason = "not-numeric" Then
            VLA_Messages.RaiseMsg "sql-aggregate-non-numeric-operand", "function", AggFunctionName(NodeAggKind(failNode))
        Else
            VLA_Messages.RaiseMsg "sql-aggregate-empty-no-rows", "function", AggFunctionName(NodeAggKind(failNode))
        End If
    End If

    Set outColMap = VLA_Runtime.VlaDictNew()
    For gi = 1 To groupArity
        Dim gcol As Collection
        Set gcol = groupByItems.Item(gi)
        Dim gqual As String, gnm As String
        gqual = NodeColumnQualifier(gcol)
        gnm = NodeColumnName(gcol)
        If Len(gqual) > 0 Then VLA_Runtime.VlaDictSet outColMap, gqual & "." & gnm, gi
        If VLA_Runtime.VlaDictHas(outColMap, gnm) Then
            VLA_Runtime.VlaDictSet outColMap, gnm, Empty
        Else
            VLA_Runtime.VlaDictSet outColMap, gnm, gi
        End If
    Next gi
    For ak = 1 To aggList.Count
        VLA_Runtime.VlaDictSet outColMap, AggSignature(aggList.Item(ak)), groupArity + ak
    Next ak

    Set ComputeGroupedRows = groupedRows
End Function

' ---- SQL.5: ORDER BY resolution + stable multi-key sort --------------

' Resolves one ORDER BY item to WHERE its own sort value comes from - an
' ordinal is always a position in the FINAL OUTPUT row (unchanged: real
' SQL's own rule, "the Nth SELECTed thing"). A name is tried against the
' OUTPUT header list FIRST (an output alias/bare column always takes
' PRECEDENCE over a same-named source column, real SQL's own rule) and,
' only when that finds nothing, falls back to the SAME colMap WHERE/
' HAVING already resolve bare column names against - real SQL's own
' (and specifically SQLite's own) allowance to ORDER BY a column that
' was never re-selected, live-caught missing from this item's own first
' pass (a plain "ORDER BY Dept" over "SELECT Name FROM t" was wrongly
' refused - output-only resolution alone is stricter than real SQL,
' surprising for the single most common ORDER BY shape there is).
' srcKind out is ORDERSRC_OUTPUT (value lives in the already-PROJECTED
' output row, position srcPos) or ORDERSRC_SOURCE (value lives in the
' SOURCE/grouped row colMap already resolves everything else against,
' position srcPos there instead) - SqlRunJoin's own row loop below reads
' from whichever one each item's own plan says.
Private Sub ResolveOrderItem(ByVal item As Collection, ByRef outHeaders() As String, ByVal outArity As Long, _
                              ByVal colMap As Object, ByRef srcKind As Long, ByRef srcPos As Long)
    If OrderItemKind(item) = ORDER_ORDINAL Then
        Dim p As Long
        p = OrderItemOrdinal(item)
        If p > outArity Then
            VLA_Messages.RaiseMsg "sql-order-by-position-out-of-range", "position", p, "count", outArity
        End If
        srcKind = ORDERSRC_OUTPUT
        srcPos = p
        Exit Sub
    End If

    Dim nm As String
    nm = OrderItemName(item)

    Dim foundPos As Long, matchCount As Long
    foundPos = 0
    matchCount = 0
    Dim h As Long
    For h = 1 To outArity
        If StrComp(VLA_Identity.Fold(outHeaders(h)), nm, vbBinaryCompare) = 0 Then
            If matchCount = 0 Then foundPos = h
            matchCount = matchCount + 1
        End If
    Next h
    If matchCount > 1 Then VLA_Messages.RaiseMsg "sql-order-by-ambiguous-column", "column", nm
    If matchCount = 1 Then
        srcKind = ORDERSRC_OUTPUT
        srcPos = foundPos
        Exit Sub
    End If

    If VLA_Runtime.VlaDictHas(colMap, nm) Then
        Dim v As Variant
        v = VLA_Runtime.VlaDictGet(colMap, nm)
        If IsEmpty(v) Then VLA_Messages.RaiseMsg "sql-order-by-ambiguous-column", "column", nm
        srcKind = ORDERSRC_SOURCE
        srcPos = CLng(v)
        Exit Sub
    End If

    VLA_Messages.RaiseMsg "sql-order-by-unknown-column", "column", nm
End Sub

' Numeric-vs-text STRICT policy, identical to WHERE/HAVING's own
' (VLA_Relation.CompareValues, ValueIsNumericType alone) - never a
' second value-comparison policy for sorting, reused directly.
Private Function CompareOneValue(ByVal a As Variant, ByVal b As Variant) As Long
    Dim bothNum As Boolean
    bothNum = VLA_Relation.ValueIsNumericType(a) And VLA_Relation.ValueIsNumericType(b)
    If VLA_Relation.CompareValues("<", a, b, bothNum) Then
        CompareOneValue = -1
    ElseIf VLA_Relation.CompareValues(">", a, b, bothNum) Then
        CompareOneValue = 1
    Else
        CompareOneValue = 0
    End If
End Function

' Walks every ORDER BY key in order, falling through to the next one
' only on an exact tie (a real multi-key sort, not just the first key) -
' DESC flips a key's own comparison sign; returning 0 after exhausting
' every key (a full tie) is what lets the merge below stay stable.
' sortKeys(rowIndex) is a Variant() of exactly orderItems.Count already-
' EVALUATED values (SqlRunJoin's own per-row loop builds these, one per
' ORDER BY item, reading from whichever of the output/source row each
' item's own ResolveOrderItem plan says) - a uniform 1..N walk, no
' separate positions() array needed here at all, unlike an ordinary
' projected row.
Private Function CompareRowsForSort(ByRef sortKeys() As Variant, ByVal rowA As Long, ByVal rowB As Long, ByRef descs() As Boolean) As Long
    Dim keyA() As Variant, keyB() As Variant
    keyA = sortKeys(rowA)
    keyB = sortKeys(rowB)
    Dim k As Long
    For k = 1 To UBound(descs)
        Dim c As Long
        c = CompareOneValue(keyA(k), keyB(k))
        If descs(k) Then c = -c
        If c <> 0 Then
            CompareRowsForSort = c
            Exit Function
        End If
    Next k
    CompareRowsForSort = 0
End Function

' A merge sort over idx() - VBA has no native array sort, and a stable
' one (equal rows keep their own original relative order) needs writing
' by hand; merge sort is naturally stable as long as the merge step's
' own tie-break always prefers the LEFT run first (<=, below), which it
' does. Neither buf() nor sortKeys() is ever reordered in place - only
' idx() moves, this file's own SQL.5 header note has the "why"
' (BETA_ROADMAP2.md's own words: "over row indices, never swapping
' whole rows per comparison").
Private Sub StableSortIndices(ByRef idx() As Long, ByRef sortKeys() As Variant, ByRef descs() As Boolean)
    Dim n As Long
    n = UBound(idx) - LBound(idx) + 1
    If n <= 1 Then Exit Sub
    Dim tmp() As Long
    ReDim tmp(LBound(idx) To UBound(idx))
    MergeSortRange idx, tmp, sortKeys, descs, LBound(idx), UBound(idx)
End Sub

Private Sub MergeSortRange(ByRef idx() As Long, ByRef tmp() As Long, ByRef sortKeys() As Variant, _
                            ByRef descs() As Boolean, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim mid As Long
    mid = (lo + hi) \ 2
    MergeSortRange idx, tmp, sortKeys, descs, lo, mid
    MergeSortRange idx, tmp, sortKeys, descs, mid + 1, hi
    Dim i As Long, j As Long, k As Long
    i = lo
    j = mid + 1
    k = lo
    Do While i <= mid And j <= hi
        If CompareRowsForSort(sortKeys, idx(i), idx(j), descs) <= 0 Then
            tmp(k) = idx(i)
            i = i + 1
        Else
            tmp(k) = idx(j)
            j = j + 1
        End If
        k = k + 1
    Loop
    Do While i <= mid
        tmp(k) = idx(i)
        i = i + 1
        k = k + 1
    Loop
    Do While j <= hi
        tmp(k) = idx(j)
        j = j + 1
        k = k + 1
    Loop
    Dim m As Long
    For m = lo To hi
        idx(m) = tmp(m)
    Next m
End Sub

' The pure core, with no Excel dependency at all: parses queryText,
' resolves FROM plus every JOIN clause against tables (a Collection of
' MakeTableRec records - one per table this call actually has access
' to, matched by NAME, never by argument position, DATALOG's own
' predicate-naming discipline applied again), folds each JOIN into the
' accumulated row set LEFT TO RIGHT via VLA_Relation.RelJoin (SQL.3),
' evaluates WHERE (if present) over the result, and projects the SELECT
' list - a bare column reference or a computed addExpr, evaluated
' uniformly through EvalScalar (SQL.2), no separate code path for
' either shape. DISTINCT (SQL.2), when present, then drops every row
' whose OWN OUTPUT VALUES duplicate an earlier kept row. Returns a
' 3-item Collection:
' (1) headers - a Collection of Strings, in SELECT-list order (each
'     item's own AS alias, or a bare column's own original casing - or
'     every joined table's own original column names, in join order,
'     for SELECT *);
' (2) rows - a Collection of Variant() arrays, one per matched (and,
'     under DISTINCT, first-seen) row, in source order;
' (3) arity - the output's own column count (Long).
' VLA_Tests_Query.TestSql is this function's own first caller, off
' hand-built table records, never a live Range - the identical pure/
' host seam VLA_Datalog.DatalogRun already established.
Public Function SqlRunJoin(ByVal queryText As String, ByVal tables As Collection) As Collection
    Dim hasDistinct As Boolean
    Dim selectAll As Boolean
    Dim selectItems As Collection
    Dim fromName As String
    Dim joins As Collection
    Dim whereExpr As Collection
    Dim hasWhere As Boolean
    Dim hasGroupBy As Boolean
    Dim groupByItems As Collection
    Dim hasHaving As Boolean
    Dim havingExpr As Collection
    Dim hasOrderBy As Boolean
    Dim orderItems As Collection
    Dim hasLimit As Boolean
    Dim limitCount As Long
    ParseQuery queryText, hasDistinct, selectAll, selectItems, fromName, joins, whereExpr, hasWhere, _
               hasGroupBy, groupByItems, hasHaving, havingExpr, hasOrderBy, orderItems, hasLimit, limitCount

    Dim byName As Object
    Set byName = VLA_Runtime.VlaDictNew()
    Dim tv As Variant
    For Each tv In tables
        Dim trec As Collection
        Set trec = tv
        VLA_Runtime.VlaDictSet byName, TableRecName(trec), trec
    Next tv

    Dim seenTables As Object
    Set seenTables = VLA_Runtime.VlaDictNew()
    Dim colMap As Object
    Set colMap = VLA_Runtime.VlaDictNew()
    Dim outHeadersOriginal As New Collection

    If Not VLA_Runtime.VlaDictHas(byName, fromName) Then
        VLA_Messages.RaiseMsg "sql-from-table-mismatch", "from", fromName
    End If
    VLA_Runtime.VlaDictSet seenTables, fromName, True
    Dim fromRec As Collection
    Set fromRec = VLA_Runtime.VlaDictGet(byName, fromName)
    Dim accumRows As Collection
    Set accumRows = TableRecRows(fromRec)
    Dim accumArity As Long
    accumArity = TableRecColumns(fromRec).Count
    AddColumnsToMap colMap, fromName, TableRecColumns(fromRec), 0, outHeadersOriginal

    Dim jv As Variant
    For Each jv In joins
        Dim jrec As Collection
        Set jrec = jv
        Dim jname As String
        jname = JoinRecTable(jrec)
        If VLA_Runtime.VlaDictHas(seenTables, jname) Then
            VLA_Messages.RaiseMsg "sql-duplicate-join-table", "table", jname
        End If
        If Not VLA_Runtime.VlaDictHas(byName, jname) Then
            VLA_Messages.RaiseMsg "sql-join-table-not-passed", "table", jname
        End If
        VLA_Runtime.VlaDictSet seenTables, jname, True
        Dim newRec As Collection
        Set newRec = VLA_Runtime.VlaDictGet(byName, jname)
        Dim newArity As Long
        newArity = TableRecColumns(newRec).Count

        AddColumnsToMap colMap, jname, TableRecColumns(newRec), accumArity, outHeadersOriginal

        Dim equiLeft As Collection, equiRight As Collection, residualNode As Collection
        SplitOnExpr JoinRecOn(jrec), colMap, accumArity, equiLeft, equiRight, residualNode

        Dim leftCols As Variant, rightCols As Variant
        leftCols = VLA_Relation.CollToLongArray(equiLeft)
        rightCols = VLA_Relation.CollToLongArray(equiRight)

        Dim leftWrap As Collection, rightWrap As Collection
        Set leftWrap = VLA_Relation.RelWrapBag(accumArity, accumRows)
        Set rightWrap = VLA_Relation.RelWrapBag(newArity, TableRecRows(newRec))

        Dim joined As Collection
        Set joined = VLA_Relation.RelJoin(leftWrap, leftCols, rightWrap, rightCols)

        If Not residualNode Is Nothing Then
            ' `Dim filteredJoin As New Collection` here would be this
            ' codebase's own already-documented trap: this whole block
            ' sits inside the `For Each jv In joins` loop, so `As New`'s
            ' auto-instantiation (which only fires when the variable is
            ' CURRENTLY Nothing) would create a fresh Collection on the
            ' FIRST join with a residual condition only - any LATER join
            ' clause that also has one would find filteredJoin already
            ' non-Nothing (still holding the FIRST one's own filtered
            ' rows) and silently APPEND onto that same shared object
            ' instead of starting fresh, corrupting accumRows with a mix
            ' of two different joins' own rows. Explicit Set each
            ' iteration, VLA_Relation.RangeColumnNames' own `pair`
            ' (VLA_Datalog.bas's own `rec` before it) is the precedent.
            Dim filteredJoin As Collection
            Set filteredJoin = New Collection
            Dim jrow As Variant, jarr() As Variant
            For Each jrow In joined
                jarr = jrow
                If EvalBool(residualNode, colMap, jarr) Then filteredJoin.Add jarr
            Next jrow
            Set accumRows = filteredJoin
        Else
            Set accumRows = joined
        End If
        accumArity = accumArity + newArity
    Next jv

    ' SQL.4: GROUP BY / aggregates / HAVING. Whenever this query
    ' aggregates at all (ParseQuery's own identical test, repeated here
    ' since SqlRunJoin doesn't receive ParseQuery's own already-computed
    ' verdict as an output), ComputeGroupedRows turns the WHERE-filtered
    ' joined row set into ONE ROW PER GROUP plus a matching colMap keyed
    ' the same way AddColumnsToMap already keys a real table's own
    ' columns - accumRows/colMap are then REPOINTED at that grouped
    ' result, and hasWhere/whereExpr REPOINTED at hasHaving/havingExpr,
    ' so every line of code below this point (header building, the row-
    ' filter-then-project-then-DISTINCT loop) runs COMPLETELY UNCHANGED
    ' against whichever pipeline actually ran - this file's own SQL.4
    ' header note has the full design.
    If hasGroupBy Or hasHaving Or AnySelectItemHasAggregate(selectItems) Then
        Dim groupColMap As Object
        Dim groupedRows As Collection
        Set groupedRows = ComputeGroupedRows(selectItems, groupByItems, hasGroupBy, hasHaving, havingExpr, _
                                              hasWhere, whereExpr, colMap, accumRows, groupColMap)
        Set accumRows = groupedRows
        Set colMap = groupColMap
        hasWhere = hasHaving
        Set whereExpr = havingExpr
    End If

    Dim outHeaders() As String
    Dim outArity As Long
    If selectAll Then
        outArity = outHeadersOriginal.Count
        ReDim outHeaders(1 To outArity)
        Dim oi As Long
        For oi = 1 To outArity
            outHeaders(oi) = outHeadersOriginal.Item(oi)
        Next oi
    Else
        outArity = selectItems.Count
        ReDim outHeaders(1 To outArity)
        Dim i As Long
        For i = 1 To outArity
            Dim pi As Collection
            Set pi = selectItems.Item(i)
            Dim exprNode As Collection
            Set exprNode = ProjItemExpr(pi)
            ' Validated UPFRONT, once per item, before any row is ever
            ' evaluated - SQL.1's own existing "an unknown SELECT
            ' column is refused" guarantee held regardless of row
            ' count, and this preserves it for a computed expression
            ' too (EvalScalar's own NK_COLUMN check alone would only
            ' fire lazily, during row evaluation, which a zero-row
            ' table would never reach).
            ValidateScalarColumns exprNode, colMap
            If ProjItemHasAlias(pi) Then
                outHeaders(i) = ProjItemAlias(pi)
            ElseIf NodeKind(exprNode) = NK_AGG Then
                ' SQL.4: an un-aliased aggregate call - ParseQuery's own
                ' alias-requirement check exempts it (this file's own
                ' SQL.4 header note); AggDisplayText builds its own
                ' canonical default header (e.g. "COUNT(*)").
                outHeaders(i) = AggDisplayText(exprNode)
            Else
                ' Only reachable for a bare NK_COLUMN item - ParseQuery
                ' already refused any other un-aliased kind at parse
                ' time (sql-computed-column-needs-alias).
                outHeaders(i) = NodeColumnOriginal(exprNode)
            End If
        Next i
    End If

    ' SQL.5: every ORDER BY item is resolved ONCE here, UNCONDITIONALLY -
    ' before any row is ever evaluated, so a zero-row result still
    ' refuses a bad reference (ValidateScalarColumns' own upfront
    ' guarantee, extended to ORDER BY) - each item's own plan (srcKind/
    ' srcPos, ResolveOrderItem's own header note has the full resolution
    ' rule) says whether its sort VALUE will come from the projected
    ' OUTPUT row or the SOURCE/grouped row the row loop below already
    ' has colMap/arr for.
    Dim orderSrcKind() As Long, orderSrcPos() As Long, orderDesc() As Boolean
    If hasOrderBy Then
        ReDim orderSrcKind(1 To orderItems.Count)
        ReDim orderSrcPos(1 To orderItems.Count)
        ReDim orderDesc(1 To orderItems.Count)
        Dim oi3 As Long
        For oi3 = 1 To orderItems.Count
            Dim oKind3 As Long, oPos3 As Long
            ResolveOrderItem orderItems.Item(oi3), outHeaders, outArity, colMap, oKind3, oPos3
            orderSrcKind(oi3) = oKind3
            orderSrcPos(oi3) = oPos3
            orderDesc(oi3) = OrderItemDescending(orderItems.Item(oi3))
        Next oi3
    End If

    Dim buf() As Variant
    Dim used As Long, cap As Long
    used = 0
    cap = 0

    ' sortBuf runs PARALLEL to buf (1:1 by index, same includeRow gate
    ' below) - each entry is one row's own already-EVALUATED ORDER BY
    ' key tuple, read from whichever of outRow (ORDERSRC_OUTPUT) or the
    ' SOURCE row arr (ORDERSRC_SOURCE) each item's own plan says. Kept
    ' entirely separate from outRow itself (never appended onto it) so
    ' DISTINCT's own dedup (RelTryAdd, below) only ever sees real OUTPUT
    ' values, never an ORDER BY key that isn't actually selected.
    Dim sortBuf() As Variant
    Dim sortUsed As Long, sortCap As Long
    sortUsed = 0
    sortCap = 0

    Dim distinctRel As Collection
    If hasDistinct Then Set distinctRel = VLA_Relation.RelNew(outArity)

    Dim t As Variant, arr() As Variant
    For Each t In accumRows
        arr = t
        Dim keep As Boolean
        keep = True
        If hasWhere Then keep = EvalBool(whereExpr, colMap, arr)
        If keep Then
            Dim outRow() As Variant
            ReDim outRow(1 To outArity)
            Dim j As Long
            If selectAll Then
                For j = 1 To outArity
                    outRow(j) = arr(j)
                Next j
            Else
                For j = 1 To outArity
                    outRow(j) = EvalScalar(ProjItemExpr(selectItems.Item(j)), colMap, arr)
                Next j
            End If
            Dim includeRow As Boolean
            includeRow = True
            If hasDistinct Then includeRow = VLA_Relation.RelTryAdd(distinctRel, outRow)
            If includeRow Then
                BufAdd buf, used, cap, outRow
                If hasOrderBy Then
                    Dim sortRow() As Variant
                    ReDim sortRow(1 To orderItems.Count)
                    Dim ok4 As Long
                    For ok4 = 1 To orderItems.Count
                        If orderSrcKind(ok4) = ORDERSRC_OUTPUT Then
                            sortRow(ok4) = outRow(orderSrcPos(ok4))
                        Else
                            sortRow(ok4) = arr(orderSrcPos(ok4))
                        End If
                    Next ok4
                    BufAdd sortBuf, sortUsed, sortCap, sortRow
                End If
            End If
        End If
    Next t

    ' A stable multi-key sort over ROW INDICES (idx()) - neither buf()
    ' nor sortBuf() is ever reordered in place; the row-collection loop
    ' below walks idx() instead of 1..used whenever ORDER BY is present,
    ' unchanged (1..used) otherwise.
    Dim idx() As Long
    If hasOrderBy And used > 0 Then
        ReDim idx(1 To used)
        Dim ix As Long
        For ix = 1 To used
            idx(ix) = ix
        Next ix
        StableSortIndices idx, sortBuf, orderDesc
    End If

    Dim headerColl As New Collection
    Dim hi As Long
    For hi = 1 To outArity
        headerColl.Add outHeaders(hi)
    Next hi

    ' SQL.5: LIMIT - a plain row-count cutoff applied here, after
    ' whichever order (sorted or natural) the rows above ended up in;
    ' LIMIT 0 is a real, valid "zero rows," not an error.
    Dim rowColl As New Collection
    Dim ri As Long
    Dim addedCount As Long
    addedCount = 0
    For ri = 1 To used
        If hasLimit And addedCount >= limitCount Then Exit For
        Dim srcIdx As Long
        If hasOrderBy Then srcIdx = idx(ri) Else srcIdx = ri
        rowColl.Add buf(srcIdx)
        addedCount = addedCount + 1
    Next ri

    Dim outp As New Collection
    outp.Add headerColl
    outp.Add rowColl
    outp.Add outArity
    Set SqlRunJoin = outp
End Function

' SQL.1/SQL.2's own original single-table entry point, UNCHANGED
' signature and behavior - every existing pure test (TestSql) still
' calls this directly. Now a thin wrapper around SqlRunJoin's own
' generalized engine (a one-table Collection), the exact "extend
' without breaking" move VLA_Datalog.DatalogRun's own Optional
' baseRelations/headerMap already made twice.
Public Function SqlRun(ByVal queryText As String, ByVal tableName As String, _
                        ByVal columnNames As Collection, ByVal rows As Collection) As Collection
    Dim tables As New Collection
    tables.Add MakeTableRec(tableName, columnNames, rows)
    Set SqlRun = SqlRunJoin(queryText, tables)
End Function

Private Function SetOpDisplayText(ByVal opKind As Long) As String
    Select Case opKind
    Case SETOP_UNION: SetOpDisplayText = "UNION"
    Case SETOP_UNION_ALL: SetOpDisplayText = "UNION ALL"
    Case SETOP_INTERSECT: SetOpDisplayText = "INTERSECT"
    Case SETOP_EXCEPT: SetOpDisplayText = "EXCEPT"
    End Select
End Function

' SQL.6: combines two ALREADY-EVALUATED branch results (each one's own
' SqlRunJoin-shaped 3-item Collection: headers/rows/arity) with one set
' operator - column-COUNT compatibility only (never name or type),
' checked HERE, post-evaluation, since a SELECT * branch's own arity is
' only known once it has actually run against a real table (this file's
' own SQL.6 header note has the reasoning). Zero new VLA_Relation.bas
' substrate: UNION ALL is pure concatenation; UNION is dedup-the-
' concatenation (one RelNew/RelTryAdd pass over both sides in order -
' SQL.2's own DISTINCT precedent, reused directly); INTERSECT/EXCEPT
' share one branch below - a Relation built from the RIGHT side (the
' membership probe), then the LEFT side kept (INTERSECT) or dropped
' (EXCEPT) by RelContainsTuple, PLUS a SEPARATE output-dedup Relation,
' since real SQL's INTERSECT/EXCEPT both dedup their own RESULT too, not
' just filter the left side. The combined result's own headers are
' ALWAYS leftResult's - real SQL's own "the first branch names the whole
' compound result" rule, and since "left" is always closer to the
' query's own start at every level of the precedence tree, this
' recursively bubbles the TRUE first branch's headers all the way up
' regardless of tree shape (EvalCompoundNode's own header note).
Private Function CombineSetOp(ByVal opKind As Long, ByVal leftResult As Collection, ByVal rightResult As Collection) As Collection
    Dim leftArity As Long, rightArity As Long
    leftArity = leftResult.Item(3)
    rightArity = rightResult.Item(3)
    If leftArity <> rightArity Then
        VLA_Messages.RaiseMsg "sql-set-op-column-mismatch", "operator", SetOpDisplayText(opKind), "left", leftArity, "right", rightArity
    End If

    Dim leftRows As Collection, rightRows As Collection
    Set leftRows = leftResult.Item(2)
    Set rightRows = rightResult.Item(2)
    Dim outRows As New Collection
    Dim t As Variant, arr() As Variant

    Select Case opKind
    Case SETOP_UNION_ALL
        For Each t In leftRows
            outRows.Add t
        Next t
        For Each t In rightRows
            outRows.Add t
        Next t
    Case SETOP_UNION
        Dim unionRel As Collection
        Set unionRel = VLA_Relation.RelNew(leftArity)
        For Each t In leftRows
            arr = t
            If VLA_Relation.RelTryAdd(unionRel, arr) Then outRows.Add arr
        Next t
        For Each t In rightRows
            arr = t
            If VLA_Relation.RelTryAdd(unionRel, arr) Then outRows.Add arr
        Next t
    Case SETOP_INTERSECT, SETOP_EXCEPT
        Dim probeRel As Collection
        Set probeRel = VLA_Relation.RelNew(rightArity)
        For Each t In rightRows
            arr = t
            VLA_Relation.RelTryAdd probeRel, arr
        Next t
        Dim outDedup As Collection
        Set outDedup = VLA_Relation.RelNew(leftArity)
        Dim wantPresent As Boolean
        wantPresent = (opKind = SETOP_INTERSECT)
        For Each t In leftRows
            arr = t
            If VLA_Relation.RelContainsTuple(probeRel, arr) = wantPresent Then
                If VLA_Relation.RelTryAdd(outDedup, arr) Then outRows.Add arr
            End If
        Next t
    End Select

    Dim outp As New Collection
    outp.Add leftResult.Item(1)
    outp.Add outRows
    outp.Add leftArity
    Set CombineSetOp = outp
End Function

' SQL.6: recursively evaluates a compound-query tree - a LEAF re-runs
' its own already-sliced branch TEXT through the EXISTING, completely
' unchanged SqlRunJoin (this file's own SQL.6 header note has the "why":
' each branch is real, independently-parseable SQL text, and SqlRunJoin
' already IS the whole single-query pipeline this item extends without
' touching); an internal node evaluates both sides first, THEN combines
' them (CombineSetOp) - never the reverse, since the column-count check
' can only run post-evaluation.
Private Function EvalCompoundNode(ByVal node As Collection, ByVal tables As Collection) As Collection
    If CompoundNodeKind(node) = COMPOUND_LEAF Then
        Set EvalCompoundNode = SqlRunJoin(CompoundLeafText(node), tables)
    Else
        Dim leftResult As Collection, rightResult As Collection
        Set leftResult = EvalCompoundNode(CompoundOpLeft(node), tables)
        Set rightResult = EvalCompoundNode(CompoundOpRight(node), tables)
        Set EvalCompoundNode = CombineSetOp(CompoundNodeKind(node), leftResult, rightResult)
    End If
End Function

' The compound-query entry point - wraps SqlRunJoin per branch rather
' than modifying it (this file's own SQL.6 header note has the "why"):
' a query with no top-level UNION/UNION ALL/INTERSECT/EXCEPT anywhere
' degenerates to a SINGLE leaf, in which case this is byte-for-byte
' SqlRunJoin(queryText, tables) - the ORIGINAL text, not a re-sliced
' branch, so SQL.5's own ORDER BY SOURCE-column fallback survives
' completely intact for every ordinary, non-compound query (proven, not
' just asserted, by VLA_Tests_Query.TestSqlSetOps). A genuinely compound
' query evaluates its own tree (EvalCompoundNode) and then, only if
' ORDER BY/LIMIT are present, applies them to the COMBINED result -
' resolved ONLY against the combined OUTPUT header list: an empty colMap
' makes ResolveOrderItem's own SOURCE-column fallback unreachable here,
' by construction, since no single coherent source colMap survives
' combining two branches (each may have joined different tables under
' different column names) - a bare name that isn't one of the combined
' result's own output columns gets the SAME sql-order-by-unknown-column
' refusal a genuinely unresolvable name always gets, rather than a
' second, special-cased message. ORDER BY items are resolved
' UNCONDITIONALLY whenever present, even over a zero-row combined result
' - SQL.5's own "validate before any row is evaluated" discipline,
' preserved here rather than silently skipped.
Public Function SqlRunCompound(ByVal queryText As String, ByVal tables As Collection) As Collection
    Dim root As Collection
    Dim hasOrderBy As Boolean, orderItems As Collection, hasLimit As Boolean, limitCount As Long
    ParseCompoundQuery queryText, root, hasOrderBy, orderItems, hasLimit, limitCount

    If CompoundNodeKind(root) = COMPOUND_LEAF Then
        Set SqlRunCompound = SqlRunJoin(queryText, tables)
        Exit Function
    End If

    Dim combined As Collection
    Set combined = EvalCompoundNode(root, tables)

    If Not hasOrderBy And Not hasLimit Then
        Set SqlRunCompound = combined
        Exit Function
    End If

    Dim headerColl As Collection
    Set headerColl = combined.Item(1)
    Dim outArity As Long
    outArity = combined.Item(3)
    Dim outHeaders() As String
    ReDim outHeaders(1 To outArity)
    Dim hi As Long
    For hi = 1 To outArity
        outHeaders(hi) = CStr(headerColl.Item(hi))
    Next hi

    Dim orderPos() As Long, orderDesc() As Boolean
    If hasOrderBy Then
        ReDim orderPos(1 To orderItems.Count)
        ReDim orderDesc(1 To orderItems.Count)
        Dim emptyColMap As Object
        Set emptyColMap = VLA_Runtime.VlaDictNew()
        Dim oi As Long, oKind2 As Long, oPos2 As Long
        For oi = 1 To orderItems.Count
            ResolveOrderItem orderItems.Item(oi), outHeaders, outArity, emptyColMap, oKind2, oPos2
            orderPos(oi) = oPos2
            orderDesc(oi) = OrderItemDescending(orderItems.Item(oi))
        Next oi
    End If

    Dim rowsColl As Collection
    Set rowsColl = combined.Item(2)
    Dim n As Long
    n = rowsColl.Count

    Dim finalRows As New Collection
    Dim addedCount As Long
    addedCount = 0

    If hasOrderBy And n > 0 Then
        Dim buf() As Variant, sortBuf() As Variant
        ReDim buf(1 To n)
        ReDim sortBuf(1 To n)
        Dim ri As Long
        ri = 0
        Dim rv As Variant, rarr() As Variant
        For Each rv In rowsColl
            ri = ri + 1
            rarr = rv
            buf(ri) = rarr
            Dim sortRow() As Variant
            ReDim sortRow(1 To orderItems.Count)
            Dim ok As Long
            For ok = 1 To orderItems.Count
                sortRow(ok) = rarr(orderPos(ok))
            Next ok
            sortBuf(ri) = sortRow
        Next rv

        Dim idx() As Long
        ReDim idx(1 To n)
        Dim ix As Long
        For ix = 1 To n
            idx(ix) = ix
        Next ix
        StableSortIndices idx, sortBuf, orderDesc

        Dim fi As Long
        For fi = 1 To n
            If hasLimit And addedCount >= limitCount Then Exit For
            finalRows.Add buf(idx(fi))
            addedCount = addedCount + 1
        Next fi
    Else
        Dim rv2 As Variant
        For Each rv2 In rowsColl
            If hasLimit And addedCount >= limitCount Then Exit For
            finalRows.Add rv2
            addedCount = addedCount + 1
        Next rv2
    End If

    Dim outp As New Collection
    outp.Add headerColl
    outp.Add finalRows
    outp.Add outArity
    Set SqlRunCompound = outp
End Function

' ---- SQL.7: WITH / WITH RECURSIVE -------------------------------------

' Starting AT toks(openPos) - a '(' token, the CTE definition's own
' opening bracket - scans forward counting bracket depth (no SQL grammar
' awareness needed, a plain balanced-bracket scan over TK_OP tokens) and
' returns the position of the matching ')'. Reaching EOF before depth
' returns to 0 is refused by name (sql-cte-unterminated-definition)
' rather than silently consuming the rest of the query text.
Private Function FindMatchingCloseParen(ByVal toks As Collection, ByVal openPos As Long) As Long
    Dim depth As Long
    depth = 0
    Dim p As Long
    p = openPos
    Do
        Dim t As Collection
        Set t = CurTok(toks, p)
        If TokKind(t) = TK_EOF Then VLA_Messages.RaiseMsg "sql-cte-unterminated-definition"
        If TokKind(t) = TK_OP Then
            If TokText(t) = "(" Then
                depth = depth + 1
            ElseIf TokText(t) = ")" Then
                depth = depth - 1
                If depth = 0 Then
                    FindMatchingCloseParen = p
                    Exit Function
                End If
            End If
        End If
        p = p + 1
    Loop
End Function

' How many times branchText's own FROM/JOIN references cteFolded - a
' fresh, independent re-tokenize/re-parse of branchText purely to read
' its own fromName/joins back out (ParseSimpleSelect's own two grammar-
' level outputs, discarded everywhere else this function is used since
' evaluation always re-parses the SAME text again anyway, SQL.6's own
' "re-parse real text through the existing pipeline" discipline, reused
' here for a different purpose: self-reference DETECTION, not
' evaluation). Both fromName and every JoinRecTable are already folded
' at parse time (ParseSimpleSelect's own VLA_Identity.Fold calls), and
' cteFolded is folded by its own caller (SqlRunWith) - StrComp between
' two already-folded strings, ColumnNodesEqual's own precedent, reused.
Private Function LeafReferenceCount(ByVal branchText As String, ByVal cteFolded As String) As Long
    Dim toks As Collection
    Set toks = Tokenize(branchText)
    Dim pos As Long
    pos = 1
    Dim hasDistinct As Boolean, selectAll As Boolean, selectItems As Collection
    Dim fromName As String, joins As Collection, whereExpr As Collection, hasWhere As Boolean
    Dim hasGroupBy As Boolean, groupByItems As Collection, hasHaving As Boolean, havingExpr As Collection
    ParseSimpleSelect toks, pos, hasDistinct, selectAll, selectItems, fromName, joins, whereExpr, hasWhere, _
                       hasGroupBy, groupByItems, hasHaving, havingExpr

    Dim cnt As Long
    cnt = 0
    If StrComp(fromName, cteFolded, vbBinaryCompare) = 0 Then cnt = cnt + 1
    Dim jv As Variant
    For Each jv In joins
        Dim jrec As Collection
        Set jrec = jv
        If StrComp(JoinRecTable(jrec), cteFolded, vbBinaryCompare) = 0 Then cnt = cnt + 1
    Next jv
    LeafReferenceCount = cnt
End Function

' Total self-reference count across a WHOLE compound tree - used only to
' decide "does this CTE body reference itself ANYWHERE at all" (the
' sql-cte-self-reference-needs-recursive check, which must fire
' regardless of WHERE in the tree the reference sits); the NARROWER
' "is this the one canonical recursive shape" check, below, inspects
' the base/step leaves individually rather than this total.
Private Function CompoundNodeSelfRefCount(ByVal node As Collection, ByVal cteFolded As String) As Long
    If CompoundNodeKind(node) = COMPOUND_LEAF Then
        CompoundNodeSelfRefCount = LeafReferenceCount(CompoundLeafText(node), cteFolded)
    Else
        CompoundNodeSelfRefCount = CompoundNodeSelfRefCount(CompoundOpLeft(node), cteFolded) _
                                  + CompoundNodeSelfRefCount(CompoundOpRight(node), cteFolded)
    End If
End Function

' Evaluates ONE CTE's own body text (already sliced from between its
' `(...)`) against tables (real tables plus every EARLIER cte's own
' synthetic table record, already registered by SqlRunWith) - this file's
' own SQL.7 header note has the full design for both branches below.
'
' No self-reference anywhere: an ORDINARY cte - the WHOLE body text is
' handed to the EXISTING, unchanged SqlRunCompound, which already
' supports the full compoundQuery grammar including its own trailing
' ORDER BY/LIMIT (kept, honored, no special-casing needed at all).
'
' A self-reference exists: recursive, requires the RECURSIVE keyword and
' the one canonical shape this item supports (root is SETOP_UNION_ALL,
' both sides a single COMPOUND_LEAF, base referencing itself zero times,
' step referencing itself exactly once, no ORDER BY/LIMIT of its own) -
' every other shape refused by name, this file's own SQL.7 header note
' has the full catalogue. Round 1 evaluates the base once, full, through
' the EXISTING, unchanged SqlRunJoin; every later round substitutes ONLY
' the previous round's own new rows for the CTE's own name (a fresh
' MakeTableRec built fresh each round) and re-runs the step through that
' SAME unchanged SqlRunJoin - the semi-naive delta-only optimization,
' achieved entirely through SqlRunJoin's own existing table-substitution
' seam, no new evaluation machinery at all.
Private Function EvalCteBody(ByVal cteBodyText As String, ByVal cteFolded As String, _
                              ByVal hasRecursiveKw As Boolean, ByVal tables As Collection) As Collection
    Dim root As Collection
    Dim hasOrderBy As Boolean, orderItems As Collection, hasLimit As Boolean, limitCount As Long
    ParseCompoundQuery cteBodyText, root, hasOrderBy, orderItems, hasLimit, limitCount

    If CompoundNodeSelfRefCount(root, cteFolded) = 0 Then
        Set EvalCteBody = SqlRunCompound(cteBodyText, tables)
        Exit Function
    End If

    If Not hasRecursiveKw Then VLA_Messages.RaiseMsg "sql-cte-self-reference-needs-recursive", "name", cteFolded

    Dim shapeOk As Boolean
    shapeOk = (CompoundNodeKind(root) = SETOP_UNION_ALL)
    If shapeOk Then shapeOk = (CompoundNodeKind(CompoundOpLeft(root)) = COMPOUND_LEAF) And (CompoundNodeKind(CompoundOpRight(root)) = COMPOUND_LEAF)
    If Not shapeOk Then VLA_Messages.RaiseMsg "sql-cte-recursive-shape", "name", cteFolded

    Dim baseLeaf As Collection, stepLeaf As Collection
    Set baseLeaf = CompoundOpLeft(root)
    Set stepLeaf = CompoundOpRight(root)

    Dim baseRefs As Long, stepRefs As Long
    baseRefs = LeafReferenceCount(CompoundLeafText(baseLeaf), cteFolded)
    stepRefs = LeafReferenceCount(CompoundLeafText(stepLeaf), cteFolded)

    ' No "stepRefs = 0" check here - genuinely unreachable given the
    ' checks already run: shapeOk (just above) proved root is EXACTLY
    ' these two leaves, so CompoundNodeSelfRefCount(root, ...) - already
    ' confirmed > 0 by the very first check in this function - equals
    ' baseRefs + stepRefs here; baseRefs = 0 survived the check right
    ' above, so stepRefs > 0 always holds by the time this line runs.
    If baseRefs > 0 Then VLA_Messages.RaiseMsg "sql-cte-recursive-base-self-references", "name", cteFolded
    If stepRefs > 1 Then VLA_Messages.RaiseMsg "sql-cte-recursive-step-single-reference", "name", cteFolded
    If hasOrderBy Or hasLimit Then VLA_Messages.RaiseMsg "sql-cte-recursive-no-order-by-limit", "name", cteFolded

    Dim baseResult As Collection
    Set baseResult = SqlRunJoin(CompoundLeafText(baseLeaf), tables)

    Dim cteHeaders As Collection
    Set cteHeaders = baseResult.Item(1)
    Dim cteArity As Long
    cteArity = baseResult.Item(3)

    ' Dim wPair As New Collection here (textually inside this loop) would
    ' be this codebase's own already-documented trap: `As New` only
    ' auto-instantiates once, so every LATER iteration would keep
    ' Add-ing onto the SAME ever-growing pair instead of a fresh
    ' 2-item one - RangeColumnNames' own SQL.1-era live bug, repeated.
    ' Explicit Set each iteration instead.
    Dim workingColNames As New Collection
    Dim wh As Long
    For wh = 1 To cteHeaders.Count
        Dim whName As String
        whName = CStr(cteHeaders.Item(wh))
        Dim wPair As Collection
        Set wPair = New Collection
        wPair.Add VLA_Identity.Fold(whName)
        wPair.Add whName
        workingColNames.Add wPair
    Next wh

    Dim totalRows As New Collection
    Dim workingRows As Collection
    Set workingRows = baseResult.Item(2)
    Dim wv As Variant
    For Each wv In workingRows
        totalRows.Add wv
    Next wv

    Dim round As Long
    round = 1
    Do While workingRows.Count > 0
        round = round + 1
        If round > SQL_CTE_MAX_ROUNDS Then
            VLA_Messages.RaiseMsg "sql-cte-round-ceiling", "name", cteFolded, "rounds", SQL_CTE_MAX_ROUNDS
        End If

        ' Same trap, same fix, one level up: this whole block sits
        ' inside the round loop, so `As New` here would keep ADDING
        ' onto ONE ever-growing table list across every round instead
        ' of starting fresh each time (which would silently register
        ' the CTE's own name more than once per round, each addition
        ' shadowing the last via VlaDictSet's own add-or-replace - the
        ' FINAL round's own delta would win for every earlier round's
        ' own JOIN too, corrupting anything but the last round).
        ' Explicit Set each iteration, SqlRunJoin's own `filteredJoin`
        ' header note (SQL.3) is the precedent this mirrors.
        Dim roundTables As Collection
        Set roundTables = New Collection
        Dim rt As Variant
        For Each rt In tables
            roundTables.Add rt
        Next rt
        roundTables.Add MakeTableRec(cteFolded, workingColNames, workingRows)

        Dim stepResult As Collection
        Set stepResult = SqlRunJoin(CompoundLeafText(stepLeaf), roundTables)

        Dim stepArity As Long
        stepArity = stepResult.Item(3)
        If stepArity <> cteArity Then
            VLA_Messages.RaiseMsg "sql-set-op-column-mismatch", "operator", "WITH RECURSIVE", "left", cteArity, "right", stepArity
        End If

        Set workingRows = stepResult.Item(2)
        Dim nv As Variant
        For Each nv In workingRows
            totalRows.Add nv
        Next nv
    Loop

    Dim outp As New Collection
    outp.Add cteHeaders
    outp.Add totalRows
    outp.Add cteArity
    Set EvalCteBody = outp
End Function

' The WITH-clause entry point - wraps SqlRunCompound rather than
' modifying it, the SAME "wrap, don't touch" move SqlRunCompound itself
' already made for SqlRunJoin: queryText with no leading WITH keyword at
' all short-circuits to SqlRunCompound(queryText, tables) UNCHANGED,
' proven rather than just claimed by VLA_Tests_Query.TestSqlCte. Every
' CTE def's own body text is sliced from the ORIGINAL query text by
' character offset (TokStart, SQL.6's own addition) between a balanced
' `(...)` pair (FindMatchingCloseParen), evaluated in written order
' (EvalCteBody), and registered as a new synthetic table appended to an
' augmented `tables` Collection every LATER cte and the final main query
' both see - a same-named real table is simply shadowed, no special
' code needed (this file's own SQL.7 header note has the full reasoning).
Public Function SqlRunWith(ByVal queryText As String, ByVal tables As Collection) As Collection
    Dim toks As Collection
    Set toks = Tokenize(queryText)
    Dim pos As Long
    pos = 1

    If Not AtKeyword(toks, pos, "with") Then
        Set SqlRunWith = SqlRunCompound(queryText, tables)
        Exit Function
    End If
    pos = pos + 1

    Dim hasRecursiveKw As Boolean
    hasRecursiveKw = False
    If AtKeyword(toks, pos, "recursive") Then
        hasRecursiveKw = True
        pos = pos + 1
    End If

    Dim augmentedTables As New Collection
    Dim tv As Variant
    For Each tv In tables
        augmentedTables.Add tv
    Next tv

    Dim seenCteNames As Object
    Set seenCteNames = VLA_Runtime.VlaDictNew()

    Do
        Dim nameTok As Collection
        Set nameTok = CurTok(toks, pos)
        If TokKind(nameTok) <> TK_IDENT Then
            VLA_Messages.RaiseMsg "sql-expected-token", "expected", "a CTE name", "found", DescribeTok(nameTok)
        End If
        Dim cteFolded As String
        cteFolded = VLA_Identity.Fold(TokText(nameTok))
        pos = pos + 1

        If VLA_Runtime.VlaDictHas(seenCteNames, cteFolded) Then
            VLA_Messages.RaiseMsg "sql-cte-duplicate-name", "name", cteFolded
        End If
        VLA_Runtime.VlaDictSet seenCteNames, cteFolded, True

        ExpectKeyword toks, pos, "as"
        If Not AtOp(toks, pos, "(") Then
            VLA_Messages.RaiseMsg "sql-expected-token", "expected", "'('", "found", DescribeTok(CurTok(toks, pos))
        End If
        Dim openPos As Long
        openPos = pos
        Dim closePos As Long
        closePos = FindMatchingCloseParen(toks, openPos)

        Dim bodyStart As Long, bodyEnd As Long
        bodyStart = TokStart(CurTok(toks, openPos + 1))
        bodyEnd = TokStart(CurTok(toks, closePos))
        Dim cteBodyText As String
        cteBodyText = Mid$(queryText, bodyStart, bodyEnd - bodyStart)

        Dim cteResult As Collection
        Set cteResult = EvalCteBody(cteBodyText, cteFolded, hasRecursiveKw, augmentedTables)

        ' Both `As New` here would be the same trap, twice over: this
        ' whole block sits inside the outer `Do` (one iteration per CTE
        ' def), and the inner loop below it runs once per header column
        ' - explicit Set on both, every iteration, same fix as
        ' EvalCteBody's own workingColNames build above.
        Dim cteHeaderColl As Collection
        Set cteHeaderColl = cteResult.Item(1)
        Dim cteColNames As Collection
        Set cteColNames = New Collection
        Dim hi2 As Long
        For hi2 = 1 To cteHeaderColl.Count
            Dim hName As String
            hName = CStr(cteHeaderColl.Item(hi2))
            Dim hPair As Collection
            Set hPair = New Collection
            hPair.Add VLA_Identity.Fold(hName)
            hPair.Add hName
            cteColNames.Add hPair
        Next hi2
        augmentedTables.Add MakeTableRec(cteFolded, cteColNames, cteResult.Item(2))

        pos = closePos + 1
        If AtOp(toks, pos, ",") Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop

    Dim mainStart As Long
    mainStart = TokStart(CurTok(toks, pos))
    Dim mainText As String
    mainText = Mid$(queryText, mainStart)
    Set SqlRunWith = SqlRunCompound(mainText, augmentedTables)
End Function

' result's own shape, header row first, matching SELECTROWS'/DATALOG's
' own frozen contract - always a header row, even over zero matching
' rows (arity is always >= 1 here - the parser's own select-list loop
' can never produce zero columns, ParseQuery's own note has the reason -
' so this never hits the inverted-bounds ReDim hazard VLA_Relation.bas's
' own RelUnit/SourceToArray headers document; SQL.1 has no (headless)
' directive to need the same care DATALOG's own zero-row-headless case
' took).
Private Function BuildSpilledArray(ByVal result As Collection) As Variant
    Dim headers As Collection
    Set headers = result.Item(1)
    Dim rows As Collection
    Set rows = result.Item(2)
    Dim arity As Long
    arity = result.Item(3)
    Dim n As Long
    n = rows.Count
    Dim out() As Variant
    ReDim out(1 To n + 1, 1 To arity)
    Dim i As Long
    For i = 1 To arity
        out(1, i) = headers.Item(i)
    Next i
    Dim r As Long
    r = 1
    Dim rowV As Variant, rowArr() As Variant
    For Each rowV In rows
        rowArr = rowV
        r = r + 1
        For i = 1 To arity
            out(r, i) = rowArr(i)
        Next i
    Next rowV
    BuildSpilledArray = out
End Function

' Thin wrapper over VLA_Relation.TableArgResolve, turning its message-
' agnostic reason code back into SQL's own wording - VLA_Datalog.
' TableArgName's own identical wrapper is the precedent this mirrors.
Private Function SqlTableName(ByVal v As Variant) As String
    Dim ok As Boolean, reason As String
    SqlTableName = VLA_Relation.TableArgResolve(v, ok, reason)
    If ok Then Exit Function
    Select Case reason
    Case "not-a-range"
        VLA_Messages.RaiseMsg "sql-table-not-a-range"
    Case "noncontiguous"
        VLA_Messages.RaiseMsg "sql-table-noncontiguous-columns"
    Case "needs-a-name"
        VLA_Messages.RaiseMsg "sql-table-needs-a-name"
    End Select
End Function

' Thin wrapper over VLA_Relation.RangeColumnNames - ok=False there means
' "no real ListObject", which for SQL specifically (unlike DATALOG, that
' never needed named columns at all) is a hard requirement, not an
' alternate path.
Private Function SqlColumnNames(ByVal v As Variant) As Collection
    Dim ok As Boolean
    Set SqlColumnNames = VLA_Relation.RangeColumnNames(v, ok)
    If Not ok Then VLA_Messages.RaiseMsg "sql-table-needs-real-table"
End Function

' The worksheet-facing entry point - =SQL(query, table1, table2, ...).
' Deliberately ParamArray-shaped from SQL.1 on (Public Function
' SQL(query, ParamArray tables())), not the bare SQL(table, query)
' two-argument shape this item's own roadmap title used as a mnemonic:
' VBA requires a ParamArray to be a function's own LAST parameter (a
' hard language rule, DATALOG's own header already relies on this) -
' this signature was always the one shape that could grow to SQL.3's
' own multiple tables without ever breaking a formula written against
' SQL.1, mirroring DATALOG's own (rules, table1, table2, ...) shape
' exactly, and SQL.3 is that growth arriving.
'
' SQL.3: at least one table argument is required now
' (sql-needs-at-least-one-table, retiring - never reusing, SD-9 -
' SQL.1's own sql-needs-exactly-one-table, which this file no longer
' raises); every table FROM/JOIN actually names must correspond to one
' of the PASSED arguments, matched by its own real name (SqlRunJoin's
' own byName lookup), never by argument position - a query naming
' fewer tables than were passed is fine (DATALOG's own precedent: an
' unreferenced table argument is simply unused, never an error).
'
' Errors return as readable TEXT in the cell ("#SQL! ...") rather than
' Excel's own opaque #VALUE! - DATALOG's own precedent, reused rather
' than re-derived for a second engine's own worksheet function.
'
' SQL.6: calls SqlRunCompound rather than SqlRunJoin directly - safe
' precisely because SqlRunCompound's own short-circuit (no top-level set
' operator anywhere in queryText) degenerates to calling SqlRunJoin with
' this SAME queryText, unchanged - so every existing =SQL(...) formula
' keeps behaving identically, and a query actually using UNION/UNION
' ALL/INTERSECT/EXCEPT now works, through the SAME function.
'
' SQL.7: calls SqlRunWith now instead - the identical degenerate-to-the-
' existing-function safety argument, one layer further out: no leading
' WITH keyword at all short-circuits straight to SqlRunCompound(
' queryText, tables), unchanged, so this swap changes nothing for every
' query written before this item shipped.
Public Function SQL(ByVal queryText As String, ParamArray tables() As Variant) As Variant
    On Error GoTo fail
    Dim tblCount As Long
    tblCount = UBound(tables) - LBound(tables) + 1
    If tblCount < 1 Then VLA_Messages.RaiseMsg "sql-needs-at-least-one-table", "count", tblCount

    Dim tableRecs As New Collection
    Dim i As Long
    For i = LBound(tables) To UBound(tables)
        Dim tableName As String
        tableName = SqlTableName(tables(i))
        Dim colNames As Collection
        Set colNames = SqlColumnNames(tables(i))
        Dim rows As Collection
        Set rows = VLA_Relation.RangeToRows(tables(i))
        tableRecs.Add MakeTableRec(tableName, colNames, rows)
    Next i

    Dim result As Collection
    Set result = SqlRunWith(queryText, tableRecs)
    SQL = BuildSpilledArray(result)
    Exit Function
fail:
    SQL = "#SQL! " & Err.Description
End Function
