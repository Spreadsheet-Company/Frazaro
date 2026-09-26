Attribute VB_Name = "VLA_OptimizeSearch"
Option Explicit
Public Const VLA_OPTIMIZE_SEARCH_VERSION As String = "OPTIMIZE.3"

' =====================================================================
'  VLA_OptimizeSearch - OPTIMIZE.3 slice 2: the search.
'
'  WHAT IT IS. One decision at a time over integer atoms, with what each
'  decision forces propagated before the next is made, on an explicit
'  trail, iteratively. VLA_Optimize grounds a program into the problem
'  this module takes - atoms numbered 1..n, CLAUSES over them (a
'  constraint made ground) and COUNTERS (a choice form's "exactly",
'  "at least", "at most" and "between", per group) - and this module
'  knows nothing about rules, Tables or text. It answers one of four
'  things: a world that satisfies every clause and counter; proof that
'  none exists; the work budget spent first; or the seconds guard fired.
'
'  WHICH WORLD, and it is the owner's call (OPTIMIZE.3 fork 6,
'  2026-09-24): atoms are decided in NUMBER ORDER, the lowest undecided
'  first, and each is tried TRUE before FALSE. VLA_Optimize numbers the
'  atoms in the Tables' own order, so the answer is the first valid
'  world in that order - lexicographically first, TRUE before FALSE -
'  which explains in one sentence and is the same answer every later
'  improvement that only removes dead ends must keep (learning,
'  backjumping, root probing, components). Propagation cannot change
'  which world is first: it only ever assigns a value every world
'  extending the current one must share, so it removes dead ends and
'  never a solution.
'
'  PROPAGATION, and what "a counter that forces" means. A clause is a
'  list of literals, at least one of which must hold (+a: atom a is
'  chosen; -a: it is not). When every literal but one is false, the last
'  one is forced; when all are false, that is a conflict. A counter says
'  at least `lo` and at most `hi` of its members are chosen: once `hi`
'  are chosen, every other member is forced out; once only `lo` can
'  still be chosen, every one of them is forced in. That is OPTIMIZE.0's
'  second trap closed - a counter checked only when complete is
'  enumeration with extra steps.
'
'  THE BOOKKEEPING. Each clause and counter keeps how many of its
'  members have been PROCESSED true and false. Those counts only decide
'  WHEN to look; what the look decides is read from the atoms' actual
'  values, so an assignment still waiting on the trail can never be
'  missed or counted twice. An atom's counts are all updated before any
'  of its consequences is examined, so a conflict found half way through
'  still leaves every count exactly undoable.
'
'  BACKTRACKING IS CHRONOLOGICAL. A decision that led to a conflict is
'  undone with everything after it and its atom set FALSE one level
'  down, where it is undone in turn if that level fails. No clause is
'  learned (OPTIMIZE.9, which has its trigger in the ladder). A conflict
'  at level 0 is proof that no world exists.
'
'  A CONTRADICTION FOUND BEFORE ANY CHOICE IS MADE IS EXPLAINED. When the
'  program's own requirements collide before the first decision
'  (optimize-quote-conflict: the customer wants the carbon frame and the
'  child seat, and the seat is not rated for the frame), every forced
'  value has a reason - the clause or counter that forced it - and the
'  reasons are walked back from the conflict to the rules that started
'  it. Those are the rules the answer names. After a decision the
'  question "why" has no one-line answer, and none is invented.
'
'  WORK, NOT SECONDS (standing decision 2). The budget counts decisions
'  and conflicts, so a faster machine proves exactly what a slower one
'  does, and it is never passed: a dead end the budget's last decision
'  runs into ends the search there, uncounted. The seconds guard exists
'  only so that no formula can hold Excel indefinitely (OPTIMIZE.0.C);
'  it is read every 256 units of work, since reading the clock costs,
'  and when it fires the answer says so - it is the one stop that
'  depends on the machine. Timer wraps at midnight, and the elapsed time
'  is corrected for it.
'
'  THE DATA DISCIPLINE (OPTIMIZE.0.C), and it is pinned rather than
'  hoped for: tools/check_optimize_search_discipline.ps1 holds this
'  whole module to Long arrays and numbers - no object, no dictionary,
'  no collection, no late-bound call, no implicitly typed name. One
'  object per node was measured at 160 seconds a million nodes before
'  any search happened, and the regression would be SILENT: every
'  answer stays right while the search gets a hundred times slower.
'
'  NO RECURSION. A search that recursed once per decision over 4,200
'  atoms would overflow VBA's small stack (PROLOG.5.3's finding). The
'  decisions live in two arrays indexed by level.
'
'  LAYER:     Engine (VLA_Optimize's own search; nothing else calls it)
'  MAY CALL:  nothing outside this module
'  SHIPS:     add-in (VLA_Build.bas's own mods array) and the dev rig
'  PAYS INTO: OPTIMIZE.4 (counters over sums), .6 (branch-and-bound),
'             .8 (the narrowers) and .9 (learning), each of which must
'             keep the answer this module gives.
' =====================================================================

' ---- what a search can answer --------------------------------------
Public Const OPT_SEARCH_FOUND As Long = 1
Public Const OPT_SEARCH_NONE As Long = 2
Public Const OPT_SEARCH_BUDGET As Long = 3
Public Const OPT_SEARCH_GUARD As Long = 4
' A problem whose arrays do not describe a problem - an atom outside
' 1..n, a member twice in one counter. Only a caller's defect can make
' one, so VLA_Optimize raises it as a Frazaro bug.
Public Const OPT_SEARCH_BAD_PROBLEM As Long = 5

' A counter's `hi` when it has no upper bound (choose-at-least).
Public Const OPT_SEARCH_NO_MOST As Long = -1

' ---- the problem ---------------------------------------------------
'
' Clause c's literals are lits(clauseStart(c) .. clauseStart(c + 1) - 1),
' +a for "atom a is chosen" and -a for "it is not". Counter g's members
' are members(counterStart(g) .. counterStart(g + 1) - 1). The tags are
' the caller's own numbers for the rule a clause or counter came from,
' handed back when a contradiction is explained. Build one with
' OptProblemInit and the two Add procedures below, which grow the arrays
' geometrically and keep the start arrays closed.
Public Type OptSearchProblem
    nAtoms As Long
    nClauses As Long
    nLits As Long
    clauseStart() As Long
    lits() As Long
    clauseTag() As Long
    nCounters As Long
    nMembers As Long
    counterStart() As Long
    members() As Long
    counterLo() As Long
    counterHi() As Long
    counterTag() As Long
End Type

' ---- the answer ----------------------------------------------------
'
' value(a) is 1 when atom a is chosen and -1 when it is not, for a FOUND
' answer only. work = decisions + conflicts, the unit the budget counts.
' For a NONE answer reached before any decision (rootConflict), the tags
' of every clause and counter the contradiction was walked back to,
' ascending and each once.
Public Type OptSearchResult
    outcome As Long
    value() As Long
    decisions As Long
    conflicts As Long
    work As Long
    seconds As Double
    rootConflict As Boolean
    nWhyClauses As Long
    whyClauses() As Long
    nWhyCounters As Long
    whyCounters() As Long
End Type

' ---- why an atom holds the value it holds --------------------------
Private Const RSN_DECISION As Long = 0
Private Const RSN_CLAUSE As Long = 1
Private Const RSN_COUNTER As Long = 2
Private Const RSN_FLIPPED As Long = 3

' The clock is read when work is a multiple of this plus one.
Private Const GUARD_MASK As Long = 255

' ---- the working state, reset by every run --------------------------
Private mN As Long
Private mNC As Long
Private mNG As Long
Private mVal() As Long
Private mRsnKind() As Long
Private mRsnIdx() As Long
Private mTPos() As Long
Private mTrail() As Long
Private mTLen As Long
Private mQHead As Long
Private mLevel As Long
Private mDecAtom() As Long
Private mDecPos() As Long
Private mNext As Long
Private mCStart() As Long
Private mLits() As Long
Private mCLen() As Long
Private mCTrue() As Long
Private mCFalse() As Long
Private mCTag() As Long
Private mGStart() As Long
Private mMem() As Long
Private mLo() As Long
Private mHi() As Long
Private mGSize() As Long
Private mGTrue() As Long
Private mGFalse() As Long
Private mGTag() As Long
Private mPosStart() As Long
Private mPosList() As Long
Private mNegStart() As Long
Private mNegList() As Long
Private mCtrStart() As Long
Private mCtrList() As Long
Private mDecisions As Long
Private mConflicts As Long
Private mWork As Long

' =====================================================================
'  BUILDING A PROBLEM
' =====================================================================

Public Sub OptProblemInit(ByRef prob As OptSearchProblem, ByVal nAtoms As Long)
    prob.nAtoms = nAtoms
    prob.nClauses = 0
    prob.nLits = 0
    ReDim prob.clauseStart(1 To 17)
    prob.clauseStart(1) = 1
    ReDim prob.lits(1 To 16)
    ReDim prob.clauseTag(1 To 16)
    prob.nCounters = 0
    prob.nMembers = 0
    ReDim prob.counterStart(1 To 17)
    prob.counterStart(1) = 1
    ReDim prob.members(1 To 16)
    ReDim prob.counterLo(1 To 16)
    ReDim prob.counterHi(1 To 16)
    ReDim prob.counterTag(1 To 16)
End Sub

' One clause: the first nLits entries of lits(1..). The caller removes a
' repeated literal and drops a clause holding both +a and -a; an EMPTY
' clause is legal and means no world exists.
Public Sub OptProblemAddClause(ByRef prob As OptSearchProblem, ByRef lits() As Long, _
                               ByVal nLits As Long, ByVal tag As Long)
    Dim c As Long
    c = prob.nClauses + 1
    Do While c + 1 > UBound(prob.clauseStart)
        ReDim Preserve prob.clauseStart(1 To 2 * UBound(prob.clauseStart))
    Loop
    Do While c > UBound(prob.clauseTag)
        ReDim Preserve prob.clauseTag(1 To 2 * UBound(prob.clauseTag))
    Loop
    Do While prob.nLits + nLits > UBound(prob.lits)
        ReDim Preserve prob.lits(1 To 2 * UBound(prob.lits))
    Loop
    Dim i As Long
    For i = 1 To nLits
        prob.lits(prob.nLits + i) = lits(i)
    Next i
    prob.nLits = prob.nLits + nLits
    prob.clauseTag(c) = tag
    prob.nClauses = c
    prob.clauseStart(c + 1) = prob.nLits + 1
End Sub

' One counter: at least lo and at most hi (OPT_SEARCH_NO_MOST for no
' most) of the first nMem entries of mem(1..), each an atom, each once.
Public Sub OptProblemAddCounter(ByRef prob As OptSearchProblem, ByRef mem() As Long, _
                                ByVal nMem As Long, ByVal lo As Long, ByVal hi As Long, _
                                ByVal tag As Long)
    Dim g As Long
    g = prob.nCounters + 1
    Do While g + 1 > UBound(prob.counterStart)
        ReDim Preserve prob.counterStart(1 To 2 * UBound(prob.counterStart))
    Loop
    Do While g > UBound(prob.counterTag)
        ReDim Preserve prob.counterLo(1 To 2 * UBound(prob.counterTag))
        ReDim Preserve prob.counterHi(1 To 2 * UBound(prob.counterTag))
        ReDim Preserve prob.counterTag(1 To 2 * UBound(prob.counterTag))
    Loop
    Do While prob.nMembers + nMem > UBound(prob.members)
        ReDim Preserve prob.members(1 To 2 * UBound(prob.members))
    Loop
    Dim i As Long
    For i = 1 To nMem
        prob.members(prob.nMembers + i) = mem(i)
    Next i
    prob.nMembers = prob.nMembers + nMem
    prob.counterLo(g) = lo
    prob.counterHi(g) = hi
    prob.counterTag(g) = tag
    prob.nCounters = g
    prob.counterStart(g + 1) = prob.nMembers + 1
End Sub

' =====================================================================
'  THE SEARCH
' =====================================================================

' budget is the most work (decisions plus conflicts) the search may do,
' and it does no more; guardSeconds, when above zero, the most
' wall-clock time.
Public Sub OptSearchRun(ByRef prob As OptSearchProblem, ByVal budget As Long, _
                        ByVal guardSeconds As Double, ByRef res As OptSearchResult)
    res.outcome = 0
    res.decisions = 0
    res.conflicts = 0
    res.work = 0
    res.seconds = 0
    res.rootConflict = False
    res.nWhyClauses = 0
    res.nWhyCounters = 0
    ReDim res.value(0 To 0)
    ReDim res.whyClauses(0 To 0)
    ReDim res.whyCounters(0 To 0)
    If Not ProblemIsWellFormed(prob) Then
        res.outcome = OPT_SEARCH_BAD_PROBLEM
        Exit Sub
    End If
    LoadProblem prob

    Dim started As Double
    started = CDbl(Timer)
    Dim elapsed As Double
    Dim confl As Long
    confl = RootSeed()
    Do
        If confl = 0 Then confl = Propagate()
        If confl <> 0 Then
            If mLevel = 0 Then
                mConflicts = mConflicts + 1
                mWork = mWork + 1
                res.outcome = OPT_SEARCH_NONE
                If mDecisions = 0 Then
                    res.rootConflict = True
                    ExplainRootConflict confl, res
                End If
                Exit Do
            End If
            ' The budget's last unit went on the decision that led here, so
            ' the search stops at this dead end without taking it and never
            ' does more work than it was given. Slice 2 counted the dead end
            ' as well, one unit past the budget (found by slice 4's ladder).
            If mWork >= budget Then
                res.outcome = OPT_SEARCH_BUDGET
                Exit Do
            End If
            mConflicts = mConflicts + 1
            mWork = mWork + 1
            If mWork >= budget Then
                res.outcome = OPT_SEARCH_BUDGET
                Exit Do
            End If
            Backtrack
            confl = 0
        Else
            Do While mNext <= mN
                If mVal(mNext) = 0 Then Exit Do
                mNext = mNext + 1
            Loop
            If mNext > mN Then
                res.outcome = OPT_SEARCH_FOUND
                Exit Do
            End If
            If mWork >= budget Then
                res.outcome = OPT_SEARCH_BUDGET
                Exit Do
            End If
            mDecisions = mDecisions + 1
            mWork = mWork + 1
            mLevel = mLevel + 1
            mDecAtom(mLevel) = mNext
            mDecPos(mLevel) = mTLen + 1
            AssignAtom mNext, 1, RSN_DECISION, 0
        End If
        If guardSeconds > 0 Then
            If (mWork And GUARD_MASK) = 0 Then
                elapsed = CDbl(Timer) - started
                If elapsed < 0 Then elapsed = elapsed + 86400#
                If elapsed >= guardSeconds Then
                    res.outcome = OPT_SEARCH_GUARD
                    Exit Do
                End If
            End If
        End If
    Loop

    elapsed = CDbl(Timer) - started
    If elapsed < 0 Then elapsed = elapsed + 86400#
    res.seconds = elapsed
    res.decisions = mDecisions
    res.conflicts = mConflicts
    res.work = mWork
    If res.outcome = OPT_SEARCH_FOUND Then
        ReDim res.value(0 To mN)
        Dim a As Long
        For a = 1 To mN
            res.value(a) = mVal(a)
        Next a
    End If
End Sub

' Every index in range, every start array closed and non-decreasing, no
' literal 0, no counter holding an atom twice, no count below zero.
Private Function ProblemIsWellFormed(ByRef prob As OptSearchProblem) As Boolean
    Dim n As Long
    n = prob.nAtoms
    If n < 0 Or prob.nClauses < 0 Or prob.nCounters < 0 Then Exit Function
    Dim c As Long, j As Long, lit As Long
    If prob.clauseStart(1) <> 1 Then Exit Function
    For c = 1 To prob.nClauses
        If prob.clauseStart(c + 1) < prob.clauseStart(c) Then Exit Function
        For j = prob.clauseStart(c) To prob.clauseStart(c + 1) - 1
            lit = prob.lits(j)
            If lit = 0 Or lit > n Or lit < -n Then Exit Function
        Next j
    Next c
    If prob.clauseStart(prob.nClauses + 1) <> prob.nLits + 1 Then Exit Function
    Dim seen() As Long
    ReDim seen(0 To n)
    Dim g As Long, a As Long
    If prob.counterStart(1) <> 1 Then Exit Function
    For g = 1 To prob.nCounters
        If prob.counterStart(g + 1) < prob.counterStart(g) Then Exit Function
        If prob.counterLo(g) < 0 Or prob.counterHi(g) < OPT_SEARCH_NO_MOST Then Exit Function
        For j = prob.counterStart(g) To prob.counterStart(g + 1) - 1
            a = prob.members(j)
            If a < 1 Or a > n Then Exit Function
            If seen(a) = g Then Exit Function
            seen(a) = g
        Next j
    Next g
    If prob.counterStart(prob.nCounters + 1) <> prob.nMembers + 1 Then Exit Function
    ProblemIsWellFormed = True
End Function

' Copies the problem into the working arrays and builds, for every atom,
' the clauses it appears in (each sign apart) and the counters it
' belongs to - one counting pass, one prefix sum, one filling pass.
Private Sub LoadProblem(ByRef prob As OptSearchProblem)
    mN = prob.nAtoms
    mNC = prob.nClauses
    mNG = prob.nCounters
    ReDim mVal(0 To mN)
    ReDim mRsnKind(0 To mN)
    ReDim mRsnIdx(0 To mN)
    ReDim mTPos(0 To mN)
    ReDim mTrail(0 To mN)
    ReDim mDecAtom(0 To mN)
    ReDim mDecPos(0 To mN)
    mTLen = 0
    mQHead = 1
    mLevel = 0
    mNext = 1
    mDecisions = 0
    mConflicts = 0
    mWork = 0

    Dim c As Long, j As Long, lit As Long, a As Long, g As Long
    ReDim mCStart(0 To mNC + 1)
    ReDim mCLen(0 To mNC)
    ReDim mCTrue(0 To mNC)
    ReDim mCFalse(0 To mNC)
    ReDim mCTag(0 To mNC)
    ReDim mLits(0 To prob.nLits)
    For c = 1 To mNC + 1
        mCStart(c) = prob.clauseStart(c)
    Next c
    For c = 1 To mNC
        mCLen(c) = prob.clauseStart(c + 1) - prob.clauseStart(c)
        mCTag(c) = prob.clauseTag(c)
    Next c
    For j = 1 To prob.nLits
        mLits(j) = prob.lits(j)
    Next j

    ReDim mGStart(0 To mNG + 1)
    ReDim mGSize(0 To mNG)
    ReDim mGTrue(0 To mNG)
    ReDim mGFalse(0 To mNG)
    ReDim mLo(0 To mNG)
    ReDim mHi(0 To mNG)
    ReDim mGTag(0 To mNG)
    ReDim mMem(0 To prob.nMembers)
    For g = 1 To mNG + 1
        mGStart(g) = prob.counterStart(g)
    Next g
    For g = 1 To mNG
        mGSize(g) = prob.counterStart(g + 1) - prob.counterStart(g)
        mLo(g) = prob.counterLo(g)
        mHi(g) = prob.counterHi(g)
        mGTag(g) = prob.counterTag(g)
    Next g
    For j = 1 To prob.nMembers
        mMem(j) = prob.members(j)
    Next j

    ' Occurrence lists: counts first, then starts, then fill.
    Dim posCnt() As Long, negCnt() As Long, ctrCnt() As Long
    ReDim posCnt(0 To mN + 1)
    ReDim negCnt(0 To mN + 1)
    ReDim ctrCnt(0 To mN + 1)
    For j = 1 To prob.nLits
        lit = mLits(j)
        If lit > 0 Then
            posCnt(lit) = posCnt(lit) + 1
        Else
            negCnt(-lit) = negCnt(-lit) + 1
        End If
    Next j
    For j = 1 To prob.nMembers
        ctrCnt(mMem(j)) = ctrCnt(mMem(j)) + 1
    Next j
    ReDim mPosStart(0 To mN + 1)
    ReDim mNegStart(0 To mN + 1)
    ReDim mCtrStart(0 To mN + 1)
    mPosStart(1) = 1
    mNegStart(1) = 1
    mCtrStart(1) = 1
    For a = 1 To mN
        mPosStart(a + 1) = mPosStart(a) + posCnt(a)
        mNegStart(a + 1) = mNegStart(a) + negCnt(a)
        mCtrStart(a + 1) = mCtrStart(a) + ctrCnt(a)
    Next a
    ReDim mPosList(0 To mPosStart(mN + 1))
    ReDim mNegList(0 To mNegStart(mN + 1))
    ReDim mCtrList(0 To mCtrStart(mN + 1))
    Dim posAt() As Long, negAt() As Long, ctrAt() As Long
    ReDim posAt(0 To mN + 1)
    ReDim negAt(0 To mN + 1)
    ReDim ctrAt(0 To mN + 1)
    For a = 1 To mN
        posAt(a) = mPosStart(a)
        negAt(a) = mNegStart(a)
        ctrAt(a) = mCtrStart(a)
    Next a
    For c = 1 To mNC
        For j = mCStart(c) To mCStart(c + 1) - 1
            lit = mLits(j)
            If lit > 0 Then
                mPosList(posAt(lit)) = c
                posAt(lit) = posAt(lit) + 1
            Else
                mNegList(negAt(-lit)) = c
                negAt(-lit) = negAt(-lit) + 1
            End If
        Next j
    Next c
    For g = 1 To mNG
        For j = mGStart(g) To mGStart(g + 1) - 1
            a = mMem(j)
            mCtrList(ctrAt(a)) = g
            ctrAt(a) = ctrAt(a) + 1
        Next j
    Next g
End Sub

Private Sub AssignAtom(ByVal a As Long, ByVal v As Long, ByVal kind As Long, ByVal idx As Long)
    mVal(a) = v
    mRsnKind(a) = kind
    mRsnIdx(a) = idx
    mTLen = mTLen + 1
    mTrail(mTLen) = a
    mTPos(a) = mTLen
End Sub

' What the problem forces before anything is decided: an empty clause,
' a clause of one literal, a counter that can hold none (hi = 0) or must
' take every member (lo = size), and a counter no world can satisfy.
' Returns a conflict (c for clause c, -g for counter g) or 0.
Private Function RootSeed() As Long
    Dim c As Long, lit As Long, a As Long, want As Long
    For c = 1 To mNC
        If mCLen(c) = 0 Then
            RootSeed = c
            Exit Function
        End If
        If mCLen(c) = 1 Then
            lit = mLits(mCStart(c))
            If lit > 0 Then
                a = lit
                want = 1
            Else
                a = -lit
                want = -1
            End If
            If mVal(a) = 0 Then
                AssignAtom a, want, RSN_CLAUSE, c
            ElseIf mVal(a) <> want Then
                RootSeed = c
                Exit Function
            End If
        End If
    Next c
    Dim g As Long, j As Long
    For g = 1 To mNG
        If mLo(g) > mGSize(g) Then
            RootSeed = -g
            Exit Function
        End If
        If mHi(g) <> OPT_SEARCH_NO_MOST Then
            If mLo(g) > mHi(g) Then
                RootSeed = -g
                Exit Function
            End If
        End If
        If mHi(g) = 0 Then
            For j = mGStart(g) To mGStart(g + 1) - 1
                a = mMem(j)
                If mVal(a) = 0 Then
                    AssignAtom a, -1, RSN_COUNTER, g
                ElseIf mVal(a) = 1 Then
                    RootSeed = -g
                    Exit Function
                End If
            Next j
        End If
        If mLo(g) > 0 And mLo(g) = mGSize(g) Then
            For j = mGStart(g) To mGStart(g + 1) - 1
                a = mMem(j)
                If mVal(a) = 0 Then
                    AssignAtom a, 1, RSN_COUNTER, g
                ElseIf mVal(a) = -1 Then
                    RootSeed = -g
                    Exit Function
                End If
            Next j
        End If
    Next g
End Function

' Processes the trail from mQHead. Returns 0, or the conflict: c for
' clause c, -g for counter g. See the module header for why every count
' of an atom is updated before any consequence of it is examined.
Private Function Propagate() As Long
    Dim a As Long, v As Long, i As Long, c As Long, g As Long
    Do While mQHead <= mTLen
        a = mTrail(mQHead)
        mQHead = mQHead + 1
        v = mVal(a)
        For i = mPosStart(a) To mPosStart(a + 1) - 1
            c = mPosList(i)
            If v = 1 Then
                mCTrue(c) = mCTrue(c) + 1
            Else
                mCFalse(c) = mCFalse(c) + 1
            End If
        Next i
        For i = mNegStart(a) To mNegStart(a + 1) - 1
            c = mNegList(i)
            If v = -1 Then
                mCTrue(c) = mCTrue(c) + 1
            Else
                mCFalse(c) = mCFalse(c) + 1
            End If
        Next i
        For i = mCtrStart(a) To mCtrStart(a + 1) - 1
            g = mCtrList(i)
            If v = 1 Then
                mGTrue(g) = mGTrue(g) + 1
            Else
                mGFalse(g) = mGFalse(g) + 1
            End If
        Next i

        If v = -1 Then
            For i = mPosStart(a) To mPosStart(a + 1) - 1
                c = mPosList(i)
                If mCTrue(c) = 0 Then
                    If mCFalse(c) >= mCLen(c) - 1 Then
                        If Not ClauseForce(c) Then
                            Propagate = c
                            Exit Function
                        End If
                    End If
                End If
            Next i
        Else
            For i = mNegStart(a) To mNegStart(a + 1) - 1
                c = mNegList(i)
                If mCTrue(c) = 0 Then
                    If mCFalse(c) >= mCLen(c) - 1 Then
                        If Not ClauseForce(c) Then
                            Propagate = c
                            Exit Function
                        End If
                    End If
                End If
            Next i
        End If
        For i = mCtrStart(a) To mCtrStart(a + 1) - 1
            g = mCtrList(i)
            If v = 1 Then
                If mHi(g) <> OPT_SEARCH_NO_MOST Then
                    If mGTrue(g) >= mHi(g) Then
                        If Not CounterForceHigh(g) Then
                            Propagate = -g
                            Exit Function
                        End If
                    End If
                End If
            Else
                If mGSize(g) - mGFalse(g) <= mLo(g) Then
                    If Not CounterForceLow(g) Then
                        Propagate = -g
                        Exit Function
                    End If
                End If
            End If
        Next i
    Loop
End Function

' Clause c, read from the atoms' actual values: satisfied, unit (and its
' last literal forced), or - with every literal false - a conflict.
Private Function ClauseForce(ByVal c As Long) As Boolean
    Dim j As Long, lit As Long, a As Long, x As Long
    Dim nOpen As Long, openLit As Long
    For j = mCStart(c) To mCStart(c + 1) - 1
        lit = mLits(j)
        If lit > 0 Then
            a = lit
        Else
            a = -lit
        End If
        x = mVal(a)
        If x = 0 Then
            nOpen = nOpen + 1
            openLit = lit
        ElseIf (x = 1 And lit > 0) Or (x = -1 And lit < 0) Then
            ClauseForce = True
            Exit Function
        End If
    Next j
    If nOpen = 0 Then Exit Function
    If nOpen = 1 Then
        If openLit > 0 Then
            AssignAtom openLit, 1, RSN_CLAUSE, c
        Else
            AssignAtom -openLit, -1, RSN_CLAUSE, c
        End If
    End If
    ClauseForce = True
End Function

' Counter g with hi members already chosen: every other member is forced
' out. More than hi chosen is a conflict.
Private Function CounterForceHigh(ByVal g As Long) As Boolean
    Dim j As Long, a As Long, nTrue As Long
    For j = mGStart(g) To mGStart(g + 1) - 1
        If mVal(mMem(j)) = 1 Then nTrue = nTrue + 1
    Next j
    If nTrue > mHi(g) Then Exit Function
    If nTrue = mHi(g) Then
        For j = mGStart(g) To mGStart(g + 1) - 1
            a = mMem(j)
            If mVal(a) = 0 Then AssignAtom a, -1, RSN_COUNTER, g
        Next j
    End If
    CounterForceHigh = True
End Function

' Counter g with only lo members still able to be chosen: every one of
' them is forced in. Fewer than lo able is a conflict.
Private Function CounterForceLow(ByVal g As Long) As Boolean
    Dim j As Long, a As Long, nAble As Long
    For j = mGStart(g) To mGStart(g + 1) - 1
        If mVal(mMem(j)) <> -1 Then nAble = nAble + 1
    Next j
    If nAble < mLo(g) Then Exit Function
    If nAble = mLo(g) Then
        For j = mGStart(g) To mGStart(g + 1) - 1
            a = mMem(j)
            If mVal(a) = 0 Then AssignAtom a, 1, RSN_COUNTER, g
        Next j
    End If
    CounterForceLow = True
End Function

' The newest decision failed: undo it with everything after it, and set
' its atom FALSE one level down, where it stays until that level fails.
' Every atom below the decision's own was assigned before it was made -
' it was the lowest undecided - so the next one to decide is above it.
Private Sub Backtrack()
    Dim d As Long
    d = mDecAtom(mLevel)
    UndoTo mDecPos(mLevel) - 1
    mLevel = mLevel - 1
    AssignAtom d, -1, RSN_FLIPPED, 0
    mNext = d + 1
End Sub

' Unassigns every atom after trail position keep. An atom already
' processed has its counts taken back; one still waiting never added any.
Private Sub UndoTo(ByVal keep As Long)
    Dim p As Long, a As Long, v As Long, i As Long, c As Long, g As Long
    For p = mTLen To keep + 1 Step -1
        a = mTrail(p)
        If p < mQHead Then
            v = mVal(a)
            For i = mPosStart(a) To mPosStart(a + 1) - 1
                c = mPosList(i)
                If v = 1 Then
                    mCTrue(c) = mCTrue(c) - 1
                Else
                    mCFalse(c) = mCFalse(c) - 1
                End If
            Next i
            For i = mNegStart(a) To mNegStart(a + 1) - 1
                c = mNegList(i)
                If v = -1 Then
                    mCTrue(c) = mCTrue(c) - 1
                Else
                    mCFalse(c) = mCFalse(c) - 1
                End If
            Next i
            For i = mCtrStart(a) To mCtrStart(a + 1) - 1
                g = mCtrList(i)
                If v = 1 Then
                    mGTrue(g) = mGTrue(g) - 1
                Else
                    mGFalse(g) = mGFalse(g) - 1
                End If
            Next i
        End If
        mVal(a) = 0
    Next p
    mTLen = keep
    If mQHead > keep + 1 Then mQHead = keep + 1
End Sub

' A contradiction reached with no decision made: walks every forced
' value back through the clause or counter that forced it, from the
' conflict to the rules that started it, and hands back their tags.
' A counter forced a member out because others were in (or in because
' others were out), so only those assigned before the member are its
' reasons - which keeps the rules named to the ones that mattered.
Private Sub ExplainRootConflict(ByVal confl As Long, ByRef res As OptSearchResult)
    Dim markC() As Long, markG() As Long, seenA() As Long, queue() As Long
    ReDim markC(0 To mNC)
    ReDim markG(0 To mNG)
    ReDim seenA(0 To mN)
    ReDim queue(0 To mN + 1)
    Dim qIn As Long, qOut As Long
    Dim j As Long, lit As Long, a As Long, x As Long, g As Long, c As Long, nTrue As Long
    If confl > 0 Then
        markC(confl) = 1
        For j = mCStart(confl) To mCStart(confl + 1) - 1
            lit = mLits(j)
            If lit > 0 Then a = lit Else a = -lit
            If seenA(a) = 0 Then
                seenA(a) = 1
                qIn = qIn + 1
                queue(qIn) = a
            End If
        Next j
    Else
        g = -confl
        markG(g) = 1
        For j = mGStart(g) To mGStart(g + 1) - 1
            If mVal(mMem(j)) = 1 Then nTrue = nTrue + 1
        Next j
        For j = mGStart(g) To mGStart(g + 1) - 1
            a = mMem(j)
            x = mVal(a)
            If (mHi(g) <> OPT_SEARCH_NO_MOST And nTrue > mHi(g) And x = 1) Or _
               ((mHi(g) = OPT_SEARCH_NO_MOST Or nTrue <= mHi(g)) And x = -1) Then
                If seenA(a) = 0 Then
                    seenA(a) = 1
                    qIn = qIn + 1
                    queue(qIn) = a
                End If
            End If
        Next j
    End If
    Do While qOut < qIn
        qOut = qOut + 1
        x = queue(qOut)
        If mRsnKind(x) = RSN_CLAUSE Then
            c = mRsnIdx(x)
            markC(c) = 1
            For j = mCStart(c) To mCStart(c + 1) - 1
                lit = mLits(j)
                If lit > 0 Then a = lit Else a = -lit
                If seenA(a) = 0 And mVal(a) <> 0 Then
                    seenA(a) = 1
                    qIn = qIn + 1
                    queue(qIn) = a
                End If
            Next j
        ElseIf mRsnKind(x) = RSN_COUNTER Then
            g = mRsnIdx(x)
            markG(g) = 1
            For j = mGStart(g) To mGStart(g + 1) - 1
                a = mMem(j)
                If seenA(a) = 0 And mVal(a) <> 0 Then
                    If mTPos(a) < mTPos(x) And mVal(a) = -mVal(x) Then
                        seenA(a) = 1
                        qIn = qIn + 1
                        queue(qIn) = a
                    End If
                End If
            Next j
        End If
    Loop
    CollectTags markC, mNC, mCTag, res.whyClauses, res.nWhyClauses
    CollectTags markG, mNG, mGTag, res.whyCounters, res.nWhyCounters
End Sub

' The tags of the marked entries, ascending and each once.
Private Sub CollectTags(ByRef marks() As Long, ByVal n As Long, ByRef tags() As Long, _
                        ByRef outp() As Long, ByRef nOut As Long)
    ReDim outp(0 To n)
    nOut = 0
    Dim i As Long, k As Long, t As Long, dup As Boolean
    For i = 1 To n
        If marks(i) <> 0 Then
            t = tags(i)
            dup = False
            For k = 1 To nOut
                If outp(k) = t Then
                    dup = True
                    Exit For
                End If
            Next k
            If Not dup Then
                k = nOut
                Do While k >= 1
                    If outp(k) <= t Then Exit Do
                    outp(k + 1) = outp(k)
                    k = k - 1
                Loop
                outp(k + 1) = t
                nOut = nOut + 1
            End If
        End If
    Next i
End Sub
