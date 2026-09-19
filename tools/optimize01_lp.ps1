# optimize01_lp.ps1 - export OPTIMIZE.0.1's corpus entries as clingo programs.
#
# DEV-ONLY ORACLE, exactly as tools/optimize0_lp.ps1 (settled 2026-09-19, the
# owner's call): this script only WRITES .lp text files. It never runs clingo,
# nothing in Frazaro calls it, no clingo binary is kept in this repository, and
# SD-13 is untouched: the owner runs clingo by hand from PowerShell.
#
# One file per entry of scripts/pareto_logic.txt section 17 whose clingo check
# is "pending". The four whose check is done already came from
# tools/optimize0_lp.ps1 (the toy with -Toy; 5 x 1 tight with -People 5;
# 10 x 1 loose with -People 10 -Loose; the reference roster with -People 50
# -Weeks 4) and are not repeated here.
#
# COMPARE ONLY what the corpus key states: whether any world exists, how many
# valid worlds there are, the optimum, and how many worlds reach it. Never a
# particular world: OPTIMIZE breaks ties by the Tables' row order, and clingo
# promises no particular one of several equally good answers.
#
# Each file's header says how to run it and what to expect. Two runs for an
# entry with an objective:
#   clingo <file>.lp 0 --opt-mode=ignore    counts every valid world  (Models : N)
#   clingo <file>.lp 0 --opt-mode=optN      the optimum, and every world that
#                                           reaches it  (Optimization : c,
#                                           Optimal : K)
# An entry with no world: clingo <file>.lp  (UNSATISFIABLE).
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\optimize01_lp.ps1
#   ... -Only optimize-seat-wedding     one entry
#   ... -OutDir C:\somewhere            default: $env:TEMP
#   ... -List                           print the commands and expectations only

param(
    [string]$Only = '',
    [string]$OutDir = $env:TEMP,
    [switch]$List
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

$programs = New-Object 'System.Collections.Generic.List[object]'

$programs.Add(@{ Id = 'optimize-sod-close'; Worlds = 84; Opt = '0'; Optimal = 36; Text = @'
person(ann;bob;cy;di). senior(ann;bob). task(t1;t2;t3).
% every task gets exactly one preparer and exactly one reviewer, from the staff
1 { prep(T,P) : person(P) } 1 :- task(T).
1 { rev(T,P) : person(P) } 1 :- task(T).
% nobody reviews a task they prepared; every reviewer is senior
:- prep(T,P), rev(T,P).
:- rev(T,P), not senior(P).
% nobody holds more than two roles in the month
:- person(P), #count { T,prep : prep(T,P) ; T,rev : rev(T,P) } > 2.
% as few preparing roles as possible go to seniors
#minimize { 1,T : prep(T,P), senior(P) }.
#show prep/2. #show rev/2.
'@ })

$programs.Add(@{ Id = 'optimize-sod-check'; Worlds = 0; Opt = $null; Optimal = 0; Check = $true; Text = @'
person(ann;bob;cy;di). senior(ann;bob). task(t1;t2;t3).
% this month's assignment: no choice at all
prep(t1,ann). rev(t1,bob). prep(t2,cy). rev(t2,cy). prep(t3,bob). rev(t3,ann).
viol(self_review,T) :- prep(T,P), rev(T,P).
viol(not_senior,T) :- rev(T,P), not senior(P).
viol(over_two_roles,P) :- person(P), #count { T,prep : prep(T,P) ; T,rev : rev(T,P) } > 2.
% strict=1 (the default): a violation is a constraint, so no world exists.
% Run with -c strict=0 to list the violating rows instead.
#const strict=1.
:- viol(K,X), strict = 1.
#show viol/2.
'@ })

$programs.Add(@{ Id = 'optimize-sod-short'; Worlds = 0; Opt = $null; Optimal = 0; Text = @'
person(ann;bob;cy). task(t1;t2;t3;t4).
1 { prep(T,P) : person(P) } 1 :- task(T).
1 { rev(T,P) : person(P) } 1 :- task(T).
:- prep(T,P), rev(T,P).
:- person(P), #count { T,prep : prep(T,P) ; T,rev : rev(T,P) } > 2.
#show prep/2. #show rev/2.
'@ })

$programs.Add(@{ Id = 'optimize-audit-independence'; Worlds = 90; Opt = '0'; Optimal = 9; Text = @'
staff(ann;bob;cy;di;ed). client(acme;beta;cafe).
interest(ann,acme). interest(cy,acme). interest(bob,beta). interest(di,cafe).
% every engagement gets exactly two of the staff
2 { on(C,S) : staff(S) } 2 :- client(C).
% nobody audits a client they hold an interest in; nobody on more than two
:- on(C,S), interest(S,C).
:- staff(S), #count { C : on(C,S) } > 2.
% as few engagements as possible have Ed on them
#minimize { 1,C : on(C,ed) }.
#show on/2.
'@ })

$programs.Add(@{ Id = 'optimize-duty-rotation'; Worlds = 4; Opt = $null; Optimal = 0; Text = @'
person(ann;bob;cy). duty(cash;payables;payroll). month(1..3).
last(ann,cash). last(bob,payables). last(cy,payroll).
% every month, every person gets exactly one duty, and every duty one person
1 { does(M,P,D) : duty(D) } 1 :- month(M), person(P).
:- month(M), duty(D), #count { P : does(M,P,D) } != 1.
% nobody does last month's duty in month 1, or the same duty twice in the quarter
:- does(1,P,D), last(P,D).
:- does(M1,P,D), does(M2,P,D), M1 < M2.
#show does/3.
'@ })

$programs.Add(@{ Id = 'optimize-roster-apart'; Worlds = 60; Opt = '0'; Optimal = 34; Text = @'
person(brianna;tyler;uma;vic). shift(fri;sat;sun).
request(brianna,sat).
% every shift gets exactly two of the people
2 { on(S,P) : person(P) } 2 :- shift(S).
% Brianna and Tyler never work the same shift; nobody works more than two
:- on(S,brianna), on(S,tyler).
:- person(P), #count { S : on(S,P) } > 2.
% as few requests broken as possible
#minimize { 1,P,S : on(S,P), request(P,S) }.
#show on/2.
'@ })

$programs.Add(@{ Id = 'optimize-roster-kept'; Worlds = 16; Opt = '2'; Optimal = 1; Text = @'
person(ann;bob;cy). day(1..4).   % Mon, Tue, Wed, Thu
kept(1,ann). kept(2,bob). kept(3,ann). kept(4,bob).
leave(bob,4).
% every shift gets exactly one of the people
1 { on(D,P) : person(P) } 1 :- day(D).
% nobody on leave, nobody two days in a row, nobody more than two shifts
:- on(D,P), leave(P,D).
:- on(D,P), on(D+1,P).
:- person(P), #count { D : on(D,P) } > 2.
% each difference from the kept roster costs one: an added and a dropped assignment
#minimize { 1,D,P,add : on(D,P), not kept(D,P) ; 1,D,P,drop : kept(D,P), not on(D,P) }.
#show on/2.
'@ })

$programs.Add(@{ Id = 'optimize-seat-dinner'; Worlds = 2; Opt = '0'; Optimal = 1; Text = @'
guest(ann;bob;cy;di;ed;fay). table(head;second).
friends(ed,fay). friends(cy,ed).
% every guest gets exactly one table; every table seats exactly three
1 { at(G,T) : table(T) } 1 :- guest(G).
:- table(T), #count { G : at(G,T) } != 3.
% Ann at the head table, with Bob; Cy and Di apart
:- not at(ann,head).
:- at(ann,T), not at(bob,T).
:- at(cy,T), at(di,T).
% as few pairs of friends split up as possible
split(A,B) :- friends(A,B), at(A,T), not at(B,T).
#minimize { 1,A,B : split(A,B) }.
#show at/2.
'@ })

$programs.Add(@{ Id = 'optimize-seat-wedding'; Worlds = 20; Opt = '1'; Optimal = 8; Text = @'
guest(ann;bob;cy;di;ed;fay;gus;hal;ivy). table(head;two;three).
couple(ann,bob). couple(cy,di). apart(ed,fay). apart(ann,ed).
wish(gus,ed,2). wish(hal,ivy,1).
% every guest gets exactly one table; every table seats exactly three
1 { at(G,T) : table(T) } 1 :- guest(G).
:- table(T), #count { G : at(G,T) } != 3.
% couples together, Gus at the head table, the apart pairs apart
:- couple(A,B), at(A,T), not at(B,T).
:- not at(gus,head).
:- apart(A,B), at(A,T), at(B,T).
% the least total weight of wishes broken
broken(A,B,W) :- wish(A,B,W), at(A,T), not at(B,T).
#minimize { W,A,B : broken(A,B,W) }.
#show at/2.
'@ })

$programs.Add(@{ Id = 'optimize-seat-overflow'; Worlds = 0; Opt = $null; Optimal = 0; Text = @'
guest(1..7). table(a;b).
1 { at(G,T) : table(T) } 1 :- guest(G).
:- table(T), #count { G : at(G,T) } > 3.
#show at/2.
'@ })

$programs.Add(@{ Id = 'optimize-config-laptop'; Worlds = 3; Opt = '1070'; Optimal = 1; Text = @'
item(laptop_x,laptop,900). item(laptop_y,laptop,700).
item(dock_1,dock,150). item(dock_2,dock,120).
item(monitor_4,monitor,300). item(monitor_5,monitor,250).
fits(laptop_x,dock_1). fits(laptop_y,dock_1). fits(laptop_y,dock_2).
fits(dock_1,monitor_4). fits(dock_1,monitor_5). fits(dock_2,monitor_5).
kind(K) :- item(_,K,_).
% exactly one item of each kind
1 { pick(I) : item(I,K,_) } 1 :- kind(K).
% the laptop fits the dock, and the dock fits the monitor
:- pick(L), item(L,laptop,_), pick(D), item(D,dock,_), not fits(L,D).
:- pick(D), item(D,dock,_), pick(M), item(M,monitor,_), not fits(D,M).
% the total price is at most 1,200; the lowest total price
:- #sum { P,I : pick(I), item(I,_,P) } > 1200.
#minimize { P,I : pick(I), item(I,_,P) }.
#show pick/1.
'@ })

$programs.Add(@{ Id = 'optimize-quote-bike'; Worlds = 3; Opt = '-335'; Optimal = 1; Text = @'
item(alloy_frame,frame,800,200). item(carbon_frame,frame,1500,500).
item(standard_wheels,wheels,300,60). item(carbon_wheels,wheels,700,250).
extra(child_seat,150,40). extra(lights,80,30). extra(rack,100,35).
kind(K) :- item(_,K,_,_).
price(I,P) :- item(I,_,P,_). price(E,P) :- extra(E,P,_).
margin(I,M) :- item(I,_,_,M). margin(E,M) :- extra(E,_,M).
% exactly one frame and one set of wheels; each extra in the quote or not
1 { pick(I) : item(I,K,_,_) } 1 :- kind(K).
{ pick(E) : extra(E,_,_) }.
% carbon wheels need the carbon frame; the child seat is not rated for it;
% the quote includes the child seat; the total price is at most 1,350
:- pick(carbon_wheels), not pick(carbon_frame).
:- pick(child_seat), pick(carbon_frame).
:- not pick(child_seat).
:- #sum { P,I : pick(I), price(I,P) } > 1350.
% the highest total margin (clingo prints a maximum negated)
#maximize { M,I : pick(I), margin(I,M) }.
#show pick/1.
'@ })

$programs.Add(@{ Id = 'optimize-quote-conflict'; Worlds = 0; Opt = $null; Optimal = 0; Text = @'
item(alloy_frame,frame,800,200). item(carbon_frame,frame,1500,500).
item(standard_wheels,wheels,300,60). item(carbon_wheels,wheels,700,250).
extra(child_seat,150,40). extra(lights,80,30). extra(rack,100,35).
kind(K) :- item(_,K,_,_).
1 { pick(I) : item(I,K,_,_) } 1 :- kind(K).
{ pick(E) : extra(E,_,_) }.
:- pick(carbon_wheels), not pick(carbon_frame).
:- pick(child_seat), pick(carbon_frame).
:- not pick(child_seat).
% the customer also wants the carbon frame; there is no price limit
:- not pick(carbon_frame).
#show pick/1.
'@ })

$programs.Add(@{ Id = 'optimize-exam-slots'; Worlds = 24; Opt = '0'; Optimal = 8; Text = @'
exam(stats;ml;law;art). sitting(1..3).   % 1 Mon AM, 2 Mon PM, 3 Tue AM
takes(zoe,stats). takes(zoe,ml). takes(yan,ml). takes(yan,law). takes(xi,law). takes(xi,art).
backtoback(1,2).
% every exam gets exactly one sitting
1 { in(E,S) : sitting(S) } 1 :- exam(E).
% no student has two exams in one sitting; no sitting holds more than two
:- takes(X,E1), takes(X,E2), E1 < E2, in(E1,S), in(E2,S).
:- sitting(S), #count { E : in(E,S) } > 2.
% as few students as possible with exams back to back
b2b(X) :- takes(X,E1), takes(X,E2), E1 != E2, in(E1,S1), in(E2,S2), backtoback(S1,S2).
#minimize { 1,X : b2b(X) }.
#show in/2.
'@ })

$programs.Add(@{ Id = 'optimize-class-timetable'; Worlds = 12; Opt = '0'; Optimal = 4; Text = @'
class(c7a;c7b). subject(maths;english;science). period(1..3).
% one teacher a subject, teaching both classes
% every class's every subject gets exactly one period
1 { at(C,S,P) : period(P) } 1 :- class(C), subject(S).
% a class has one lesson a period; a teacher teaches one class a period
:- class(C), period(P), #count { S : at(C,S,P) } > 1.
:- subject(S), period(P), #count { C : at(C,S,P) } > 1.
% as few maths lessons in period 1 as possible
#minimize { 1,C : at(C,maths,1) }.
#show at/3.
'@ })

$programs.Add(@{ Id = 'optimize-exam-rooms'; Worlds = 0; Opt = $null; Optimal = 0; Text = @'
exam(1..5). sitting(1..2).
1 { in(E,S) : sitting(S) } 1 :- exam(E).
:- sitting(S), #count { E : in(E,S) } > 2.
#show in/2.
'@ })

if (-not $List -and -not (Test-Path $OutDir)) { throw "no folder $OutDir" }

$written = 0
foreach ($pg in $programs) {
    if ($Only -and $pg.Id -ne $Only) { continue }
    $file = Join-Path $OutDir ($pg.Id + '.lp')
    $leaf = $pg.Id + '.lp'
    $runs = New-Object 'System.Collections.Generic.List[string]'
    if ($pg.ContainsKey('Check') -and $pg.Check) {
        $runs.Add("clingo $leaf                  expect UNSATISFIABLE")
        $runs.Add("clingo $leaf -c strict=0      expect one model: viol(self_review,t2) viol(not_senior,t2)")
    } elseif ($pg.Worlds -eq 0) {
        $runs.Add("clingo $leaf                  expect UNSATISFIABLE")
    } else {
        $runs.Add("clingo $leaf 0 --opt-mode=ignore   expect Models : $($pg.Worlds)")
        if ($null -ne $pg.Opt) { $runs.Add("clingo $leaf 0 --opt-mode=optN     expect Optimization : $($pg.Opt) and Optimal : $($pg.Optimal)") }
    }
    Write-Output "$($pg.Id)"
    foreach ($r in $runs) { Write-Output "    $r" }
    if ($List) { continue }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("% $($pg.Id) - OPTIMIZE.0.1's corpus (scripts/pareto_logic.txt section 17).")
    [void]$sb.AppendLine('% Generated by tools/optimize01_lp.ps1; do not edit by hand. Compare only what the key states.')
    foreach ($r in $runs) { [void]$sb.AppendLine("% Run:  $r") }
    [void]$sb.AppendLine('')
    [void]$sb.Append(($pg.Text -replace "`r`n", "`n"))
    [IO.File]::WriteAllText($file, $sb.ToString(), (New-Object System.Text.ASCIIEncoding))
    $written++
}
if (-not $List) { Write-Output "wrote $written file(s) to $OutDir" }
