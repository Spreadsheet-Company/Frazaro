Option Explicit

Const hot_pink = "#FF69B4"

Dim vla_problem As String

Dim vla_step As Long

Public Sub stamp(Optional row_number As Variant = 1, Optional value As Variant = "ok")
    On Error GoTo vla_fail ' vla:8
    vla_step = 52 ' vla:9
    If vlatraceon() Then ' vla:10
        Call vlatracestep(52, vla_step_text(52)) ' vla:10
    End If
    ' ---- instructions.txt:87 To stamp, with row-number of 1 and value of "ok": ----
    cells(row_number, "f") = value ' vla:13 src:90
    Exit Sub ' vla:14
vla_fail: ' vla:15
    Call vla_report_error ' vla:16
End Sub

Public Function tax(amount As Variant) As Variant
    On Error GoTo vla_fail ' vla:19
    vla_step = 102 ' vla:20
    If vlatraceon() Then ' vla:21
        Call vlatracestep(102, vla_step_text(102)) ' vla:21
    End If
    ' ---- instructions.txt:184 To tax of amount: ----
    tax = (amount * 0.08) ' vla:24 src:187
    Exit Function
    Exit Function ' vla:25
vla_fail: ' vla:26
    Call vla_report_error ' vla:27
End Function

Public Sub tidy_up()
    On Error GoTo vla_fail ' vla:30
    vla_step = 109 ' vla:31
    If vlatraceon() Then ' vla:32
        Call vlatracestep(109, vla_step_text(109)) ' vla:32
    End If
    ' ---- instructions.txt:198 To tidy-up: ----
    Call columns("a").autofit
    vla_step = 110 ' vla:36
    If vlatraceon() Then ' vla:37
        Call vlatracestep(110, vla_step_text(110)) ' vla:37
    End If
    Call columns("b").autofit
    vla_step = 111 ' vla:40
    If vlatraceon() Then ' vla:41
        Call vlatracestep(111, vla_step_text(111)) ' vla:41
    End If
    range("a1").font.color = vlacolor(hot_pink)
    Exit Sub ' vla:44
vla_fail: ' vla:45
    Call vla_report_error ' vla:46
End Sub

Public Function commission(Optional sale As Variant = 1000, Optional rate As Variant = (5 / 100)) As Variant
    On Error GoTo vla_fail ' vla:49
    vla_step = 112 ' vla:50
    If vlatraceon() Then ' vla:51
        Call vlatracestep(112, vla_step_text(112)) ' vla:51
    End If
    ' ---- instructions.txt:203 To get commission using sale of 1000 and rate of 5%: ----
    commission = (sale * rate) ' vla:54 src:206
    Exit Function
    Exit Function ' vla:55
vla_fail: ' vla:56
    Call vla_report_error ' vla:57
End Function

Public Function vat_rate() As Variant
    On Error GoTo vla_fail ' vla:60
    vla_step = 122 ' vla:61
    If vlatraceon() Then ' vla:62
        Call vlatracestep(122, vla_step_text(122)) ' vla:62
    End If
    ' ---- instructions.txt:227 To get vat-rate: ----
    vat_rate = (20 / 100) ' vla:65 src:230
    Exit Function
    Exit Function ' vla:66
vla_fail: ' vla:67
    Call vla_report_error ' vla:68
End Function

Public Sub check_text_conditions()
    On Error GoTo vla_fail ' vla:71
    vla_step = 726 ' vla:72
    If vlatraceon() Then ' vla:73
        Call vlatracestep(726, vla_step_text(726)) ' vla:73
    End If
    ' ---- instructions.txt:1231 To check-text-conditions: ----
    If (instr(1, range("aa1"), "NAN", vbtextcompare) > 0) Then ' vla:76 src:1238
        range("l17") = "contains" ' vla:76 src:1238
    End If
    vla_step = 727 ' vla:77
    If vlatraceon() Then ' vla:78
        Call vlatracestep(727, vla_step_text(727)) ' vla:78
    End If
    If (instr(1, range("aa1"), "ban", vbtextcompare) = 1) Then ' vla:80 src:1239
        range("l18") = "starts" ' vla:80 src:1239
    End If
    vla_step = 728 ' vla:81
    If vlatraceon() Then ' vla:82
        Call vlatracestep(728, vla_step_text(728)) ' vla:82
    End If
    range("l19") = "kept" ' vla:84 src:1240
    vla_step = 729 ' vla:85
    If vlatraceon() Then ' vla:86
        Call vlatracestep(729, vla_step_text(729)) ' vla:86
    End If
    If (instr(1, range("aa1"), "nan", vbtextcompare) = 0) Then ' vla:88 src:1241
        range("l19") = "wrong" ' vla:88 src:1241
    End If
    vla_step = 730 ' vla:89
    If vlatraceon() Then ' vla:90
        Call vlatracestep(730, vla_step_text(730)) ' vla:90
    End If
    If (instr(1, range("aa1"), "kiwi", vbtextcompare) = 0) Then ' vla:92 src:1242
        range("l20") = "lacks" ' vla:92 src:1242
    End If
    vla_step = 731 ' vla:93
    If vlatraceon() Then ' vla:94
        Call vlatracestep(731, vla_step_text(731)) ' vla:94
    End If
    range("l21") = "kept" ' vla:96 src:1243
    vla_step = 732 ' vla:97
    If vlatraceon() Then ' vla:98
        Call vlatracestep(732, vla_step_text(732)) ' vla:98
    End If
    If (instr(1, range("aa1"), "kiwi", vbtextcompare) > 0) Then ' vla:100 src:1244
        range("l21") = "wrong" ' vla:100 src:1244
    End If
    vla_step = 733 ' vla:101
    If vlatraceon() Then ' vla:102
        Call vlatracestep(733, vla_step_text(733)) ' vla:102
    End If
    range("l22") = "kept" ' vla:104 src:1245
    vla_step = 734 ' vla:105
    If vlatraceon() Then ' vla:106
        Call vlatracestep(734, vla_step_text(734)) ' vla:106
    End If
    If (instr(1, range("aa1"), "nan", vbtextcompare) = 1) Then ' vla:108 src:1246
        range("l22") = "wrong" ' vla:108 src:1246
    End If
    Exit Sub ' vla:109
vla_fail: ' vla:110
    Call vla_report_error ' vla:111
End Sub

Public Sub check_formulas()
    Dim gf_last As Variant ' vla:114
    Dim gf_top As Variant ' vla:115
    Dim gf_bottom As Variant ' vla:116
    Dim gf_empty As Variant ' vla:117
    Dim gf_filled As Variant ' vla:118
    On Error GoTo vla_fail ' vla:119
    vla_step = 736 ' vla:120
    If vlatraceon() Then ' vla:121
        Call vlatracestep(736, vla_step_text(736)) ' vla:121
    End If
    ' ---- instructions.txt:1250 To check-formulas: ----
    Call vlaensuresheet("gformula") ' vla:124 src:1259
    Call worksheets("gformula").activate
    vla_step = 737 ' vla:125
    If vlatraceon() Then ' vla:126
        Call vlatracestep(737, vla_step_text(737)) ' vla:126
    End If
    Call range("a1:h10").clear
    vla_step = 738 ' vla:129
    If vlatraceon() Then ' vla:130
        Call vlatracestep(738, vla_step_text(738)) ' vla:130
    End If
    range("b2") = 2 ' vla:132 src:1261
    vla_step = 739 ' vla:133
    If vlatraceon() Then ' vla:134
        Call vlatracestep(739, vla_step_text(739)) ' vla:134
    End If
    range("b3") = 3 ' vla:136 src:1262
    vla_step = 740 ' vla:137
    If vlatraceon() Then ' vla:138
        Call vlatracestep(740, vla_step_text(740)) ' vla:138
    End If
    range("b4") = 4 ' vla:140 src:1263
    vla_step = 741 ' vla:141
    If vlatraceon() Then ' vla:142
        Call vlatracestep(741, vla_step_text(741)) ' vla:142
    End If
    range("c2") = 10 ' vla:144 src:1264
    vla_step = 742 ' vla:145
    If vlatraceon() Then ' vla:146
        Call vlatracestep(742, vla_step_text(742)) ' vla:146
    End If
    range("c3") = 20 ' vla:148 src:1265
    vla_step = 743 ' vla:149
    If vlatraceon() Then ' vla:150
        Call vlatracestep(743, vla_step_text(743)) ' vla:150
    End If
    range("c4") = 30 ' vla:152 src:1266
    vla_step = 744 ' vla:153
    If vlatraceon() Then ' vla:154
        Call vlatracestep(744, vla_step_text(744)) ' vla:154
    End If
    range("d2:d4").Formula2 = "=B2*C2"
    vla_step = 745 ' vla:157
    If vlatraceon() Then ' vla:158
        Call vlatracestep(745, vla_step_text(745)) ' vla:158
    End If
    gf_last = cells(rows.count, "b").end(xlup).row ' vla:160 src:1268
    vla_step = 746 ' vla:161
    If vlatraceon() Then ' vla:162
        Call vlatracestep(746, vla_step_text(746)) ' vla:162
    End If
    If (gf_last >= 2) Then
        range(("e" & 2 & ":" & "e" & gf_last)).Formula2 = "=B2+C2"
    End If
    vla_step = 747 ' vla:165
    If vlatraceon() Then ' vla:166
        Call vlatracestep(747, vla_step_text(747)) ' vla:166
    End If
    range("f1") = "head" ' vla:168 src:1270
    vla_step = 748 ' vla:169
    If vlatraceon() Then ' vla:170
        Call vlatracestep(748, vla_step_text(748)) ' vla:170
    End If
    If (1 >= 2) Then
        range(("f" & 2 & ":" & "f" & 1)).Formula2 = "=B2-C2"
    End If
    vla_step = 749 ' vla:173
    If vlatraceon() Then ' vla:174
        Call vlatracestep(749, vla_step_text(749)) ' vla:174
    End If
    gf_top = application.worksheetfunction.max(range("d2:d4")) ' vla:176 src:1272
    vla_step = 750 ' vla:177
    If vlatraceon() Then ' vla:178
        Call vlatracestep(750, vla_step_text(750)) ' vla:178
    End If
    range("h1") = gf_top ' vla:180 src:1273
    vla_step = 751 ' vla:181
    If vlatraceon() Then ' vla:182
        Call vlatracestep(751, vla_step_text(751)) ' vla:182
    End If
    gf_bottom = application.worksheetfunction.min(range("d2:d4")) ' vla:184 src:1274
    vla_step = 752 ' vla:185
    If vlatraceon() Then ' vla:186
        Call vlatracestep(752, vla_step_text(752)) ' vla:186
    End If
    range("h2") = gf_bottom ' vla:188 src:1275
    vla_step = 753 ' vla:189
    If vlatraceon() Then ' vla:190
        Call vlatracestep(753, vla_step_text(753)) ' vla:190
    End If
    range("h3") = application.worksheetfunction.average(range("b2:b4")) ' vla:192 src:1276
    vla_step = 754 ' vla:193
    If vlatraceon() Then ' vla:194
        Call vlatracestep(754, vla_step_text(754)) ' vla:194
    End If
    range("h4") = application.worksheetfunction.max(range("c2:c4")) ' vla:196 src:1277
    vla_step = 755 ' vla:197
    If vlatraceon() Then ' vla:198
        Call vlatracestep(755, vla_step_text(755)) ' vla:198
    End If
    range("h5") = application.worksheetfunction.min(range("c2:c4")) ' vla:200 src:1278
    vla_step = 756 ' vla:201
    If vlatraceon() Then ' vla:202
        Call vlatracestep(756, vla_step_text(756)) ' vla:202
    End If
    range("b6").Formula2 = "="""""
    vla_step = 757 ' vla:205
    If vlatraceon() Then ' vla:206
        Call vlatracestep(757, vla_step_text(757)) ' vla:206
    End If
    gf_empty = application.worksheetfunction.countif(range("b2:b7"), "") ' vla:208 src:1280
    vla_step = 758 ' vla:209
    If vlatraceon() Then ' vla:210
        Call vlatracestep(758, vla_step_text(758)) ' vla:210
    End If
    range("h6") = gf_empty ' vla:212 src:1281
    vla_step = 759 ' vla:213
    If vlatraceon() Then ' vla:214
        Call vlatracestep(759, vla_step_text(759)) ' vla:214
    End If
    gf_filled = application.worksheetfunction.countif(range("b2:b7"), "<>") ' vla:216 src:1282
    vla_step = 760 ' vla:217
    If vlatraceon() Then ' vla:218
        Call vlatracestep(760, vla_step_text(760)) ' vla:218
    End If
    range("h7") = gf_filled ' vla:220 src:1283
    Exit Sub ' vla:221
vla_fail: ' vla:222
    Call vla_report_error ' vla:223
End Sub

Public Sub check_statistics()
    Dim gs_middle As Variant ' vla:226
    Dim gs_spread As Variant ' vla:227
    Dim gs_west As Variant ' vla:228
    Dim gs_none As Variant ' vla:229
    On Error GoTo vla_fail ' vla:230
    vla_step = 762 ' vla:231
    If vlatraceon() Then ' vla:232
        Call vlatracestep(762, vla_step_text(762)) ' vla:232
    End If
    ' ---- instructions.txt:1287 To check-statistics: ----
    Call vlaensuresheet("gstats") ' vla:235 src:1300
    Call worksheets("gstats").activate
    vla_step = 763 ' vla:236
    If vlatraceon() Then ' vla:237
        Call vlatracestep(763, vla_step_text(763)) ' vla:237
    End If
    Call range("a1:h10").clear
    vla_step = 764 ' vla:240
    If vlatraceon() Then ' vla:241
        Call vlatracestep(764, vla_step_text(764)) ' vla:241
    End If
    range("b2") = 2 ' vla:243 src:1302
    vla_step = 765 ' vla:244
    If vlatraceon() Then ' vla:245
        Call vlatracestep(765, vla_step_text(765)) ' vla:245
    End If
    range("b3") = 4 ' vla:247 src:1303
    vla_step = 766 ' vla:248
    If vlatraceon() Then ' vla:249
        Call vlatracestep(766, vla_step_text(766)) ' vla:249
    End If
    range("b4") = 6 ' vla:251 src:1304
    vla_step = 767 ' vla:252
    If vlatraceon() Then ' vla:253
        Call vlatracestep(767, vla_step_text(767)) ' vla:253
    End If
    range("b5") = 6 ' vla:255 src:1305
    vla_step = 768 ' vla:256
    If vlatraceon() Then ' vla:257
        Call vlatracestep(768, vla_step_text(768)) ' vla:257
    End If
    range("b6") = 7 ' vla:259 src:1306
    vla_step = 769 ' vla:260
    If vlatraceon() Then ' vla:261
        Call vlatracestep(769, vla_step_text(769)) ' vla:261
    End If
    gs_middle = application.worksheetfunction.median(range("b2:b6")) ' vla:263 src:1307
    vla_step = 770 ' vla:264
    If vlatraceon() Then ' vla:265
        Call vlatracestep(770, vla_step_text(770)) ' vla:265
    End If
    range("h1") = gs_middle ' vla:267 src:1308
    vla_step = 771 ' vla:268
    If vlatraceon() Then ' vla:269
        Call vlatracestep(771, vla_step_text(771)) ' vla:269
    End If
    gs_spread = application.worksheetfunction.stdev_s(range("b2:b6")) ' vla:271 src:1309
    vla_step = 772 ' vla:272
    If vlatraceon() Then ' vla:273
        Call vlatracestep(772, vla_step_text(772)) ' vla:273
    End If
    range("h2") = gs_spread ' vla:275 src:1310
    vla_step = 773 ' vla:276
    If vlatraceon() Then ' vla:277
        Call vlatracestep(773, vla_step_text(773)) ' vla:277
    End If
    range("h3") = application.worksheetfunction.median(range("b2:b5")) ' vla:279 src:1311
    vla_step = 774 ' vla:280
    If vlatraceon() Then ' vla:281
        Call vlatracestep(774, vla_step_text(774)) ' vla:281
    End If
    range("h4") = application.worksheetfunction.stdev_p(range("b2:b6")) ' vla:283 src:1312
    vla_step = 775 ' vla:284
    If vlatraceon() Then ' vla:285
        Call vlatracestep(775, vla_step_text(775)) ' vla:285
    End If
    range("d2") = "West" ' vla:287 src:1313
    vla_step = 776 ' vla:288
    If vlatraceon() Then ' vla:289
        Call vlatracestep(776, vla_step_text(776)) ' vla:289
    End If
    range("d3") = "East" ' vla:291 src:1314
    vla_step = 777 ' vla:292
    If vlatraceon() Then ' vla:293
        Call vlatracestep(777, vla_step_text(777)) ' vla:293
    End If
    range("d4") = "West" ' vla:295 src:1315
    vla_step = 778 ' vla:296
    If vlatraceon() Then ' vla:297
        Call vlatracestep(778, vla_step_text(778)) ' vla:297
    End If
    range("d5") = "West" ' vla:299 src:1316
    vla_step = 779 ' vla:300
    If vlatraceon() Then ' vla:301
        Call vlatracestep(779, vla_step_text(779)) ' vla:301
    End If
    range("e2") = 10 ' vla:303 src:1317
    vla_step = 780 ' vla:304
    If vlatraceon() Then ' vla:305
        Call vlatracestep(780, vla_step_text(780)) ' vla:305
    End If
    range("e3") = 20 ' vla:307 src:1318
    vla_step = 781 ' vla:308
    If vlatraceon() Then ' vla:309
        Call vlatracestep(781, vla_step_text(781)) ' vla:309
    End If
    range("e5") = 30 ' vla:311 src:1319
    vla_step = 782 ' vla:312
    If vlatraceon() Then ' vla:313
        Call vlatracestep(782, vla_step_text(782)) ' vla:313
    End If
    gs_west = application.worksheetfunction.averageif(range("d2:d5"), "West", range("e2:e5")) ' vla:315 src:1320
    vla_step = 783 ' vla:316
    If vlatraceon() Then ' vla:317
        Call vlatracestep(783, vla_step_text(783)) ' vla:317
    End If
    range("h5") = gs_west ' vla:319 src:1321
    vla_step = 784 ' vla:320
    If vlatraceon() Then ' vla:321
        Call vlatracestep(784, vla_step_text(784)) ' vla:321
    End If
    range("h6") = "kept" ' vla:323 src:1322
    vla_step = 785 ' vla:324
    If vlatraceon() Then ' vla:325
        Call vlatracestep(785, vla_step_text(785)) ' vla:325
    End If
    On Error GoTo vla_tryf_12 ' vla:327 src:1323
    vla_step = 786 ' vla:328 src:1323
    If vlatraceon() Then ' vla:329 src:1323
        Call vlatracestep(786, vla_step_text(786)) ' vla:329 src:1323
    End If
    gs_none = application.worksheetfunction.median(range("g1:g3")) ' vla:331 src:1324
    vla_step = 787 ' vla:332 src:1323
    If vlatraceon() Then ' vla:333 src:1323
        Call vlatracestep(787, vla_step_text(787)) ' vla:333 src:1323
    End If
    range("h6") = "no stop" ' vla:335 src:1325
    GoTo vla_tryd_12 ' vla:336 src:1323
vla_tryf_12: ' vla:337 src:1323
    vla_problem = err.description ' vla:338 src:1323
    Resume vla_tryr_12 ' vla:339 src:1323
vla_tryr_12: ' vla:340 src:1323
    On Error GoTo vla_fail ' vla:341 src:1323
    vla_step = 788 ' vla:342 src:1323
    If vlatraceon() Then ' vla:343 src:1323
        Call vlatracestep(788, vla_step_text(788)) ' vla:343 src:1323
    End If
    ' ---- instructions.txt:1327 If that fails: ----
    range("h7") = vla_problem ' vla:346 src:1328
vla_tryd_12: ' vla:347 src:1323
    On Error GoTo vla_fail ' vla:348 src:1323
    Exit Sub ' vla:349
vla_fail: ' vla:350
    Call vla_report_error ' vla:351
End Sub

Public Sub check_calculation()
    Dim gc_stale As Variant ' vla:354
    Dim gc_sheet As Variant ' vla:355
    Dim gc_all As Variant ' vla:356
    Dim gc_on As Variant ' vla:357
    Dim gc_auto As Variant ' vla:358
    On Error GoTo vla_fail ' vla:359
    vla_step = 790 ' vla:360
    If vlatraceon() Then ' vla:361
        Call vlatracestep(790, vla_step_text(790)) ' vla:361
    End If
    ' ---- instructions.txt:1332 To check-calculation: ----
    Call vlaensuresheet("gcalc") ' vla:364 src:1343
    Call worksheets("gcalc").activate
    vla_step = 791 ' vla:365
    If vlatraceon() Then ' vla:366
        Call vlatracestep(791, vla_step_text(791)) ' vla:366
    End If
    Call range("a1:d10").clear
    vla_step = 792 ' vla:369
    If vlatraceon() Then ' vla:370
        Call vlatracestep(792, vla_step_text(792)) ' vla:370
    End If
    range("a1") = 2 ' vla:372 src:1345
    vla_step = 793 ' vla:373
    If vlatraceon() Then ' vla:374
        Call vlatracestep(793, vla_step_text(793)) ' vla:374
    End If
    range("b1").Formula2 = "=A1*10"
    vla_step = 794 ' vla:377
    If vlatraceon() Then ' vla:378
        Call vlatracestep(794, vla_step_text(794)) ' vla:378
    End If
    application.calculation = xlcalculationmanual
    vla_step = 795 ' vla:381
    If vlatraceon() Then ' vla:382
        Call vlatracestep(795, vla_step_text(795)) ' vla:382
    End If
    range("a1") = 3 ' vla:384 src:1348
    vla_step = 796 ' vla:385
    If vlatraceon() Then ' vla:386
        Call vlatracestep(796, vla_step_text(796)) ' vla:386
    End If
    gc_stale = range("b1") ' vla:388 src:1349
    vla_step = 797 ' vla:389
    If vlatraceon() Then ' vla:390
        Call vlatracestep(797, vla_step_text(797)) ' vla:390
    End If
    range("c1") = gc_stale ' vla:392 src:1350
    vla_step = 798 ' vla:393
    If vlatraceon() Then ' vla:394
        Call vlatracestep(798, vla_step_text(798)) ' vla:394
    End If
    Call activesheet.calculate
    vla_step = 799 ' vla:397
    If vlatraceon() Then ' vla:398
        Call vlatracestep(799, vla_step_text(799)) ' vla:398
    End If
    gc_sheet = range("b1") ' vla:400 src:1352
    vla_step = 800 ' vla:401
    If vlatraceon() Then ' vla:402
        Call vlatracestep(800, vla_step_text(800)) ' vla:402
    End If
    range("c2") = gc_sheet ' vla:404 src:1353
    vla_step = 801 ' vla:405
    If vlatraceon() Then ' vla:406
        Call vlatracestep(801, vla_step_text(801)) ' vla:406
    End If
    range("a1") = 4 ' vla:408 src:1354
    vla_step = 802 ' vla:409
    If vlatraceon() Then ' vla:410
        Call vlatracestep(802, vla_step_text(802)) ' vla:410
    End If
    Call application.calculate
    vla_step = 803 ' vla:413
    If vlatraceon() Then ' vla:414
        Call vlatracestep(803, vla_step_text(803)) ' vla:414
    End If
    gc_all = range("b1") ' vla:416 src:1356
    vla_step = 804 ' vla:417
    If vlatraceon() Then ' vla:418
        Call vlatracestep(804, vla_step_text(804)) ' vla:418
    End If
    range("c3") = gc_all ' vla:420 src:1357
    vla_step = 805 ' vla:421
    If vlatraceon() Then ' vla:422
        Call vlatracestep(805, vla_step_text(805)) ' vla:422
    End If
    range("a1") = 5 ' vla:424 src:1358
    vla_step = 806 ' vla:425
    If vlatraceon() Then ' vla:426
        Call vlatracestep(806, vla_step_text(806)) ' vla:426
    End If
    application.calculation = xlcalculationautomatic
    vla_step = 807 ' vla:429
    If vlatraceon() Then ' vla:430
        Call vlatracestep(807, vla_step_text(807)) ' vla:430
    End If
    gc_on = range("b1") ' vla:432 src:1360
    vla_step = 808 ' vla:433
    If vlatraceon() Then ' vla:434
        Call vlatracestep(808, vla_step_text(808)) ' vla:434
    End If
    range("c4") = gc_on ' vla:436 src:1361
    vla_step = 809 ' vla:437
    If vlatraceon() Then ' vla:438
        Call vlatracestep(809, vla_step_text(809)) ' vla:438
    End If
    range("a1") = 6 ' vla:440 src:1362
    vla_step = 810 ' vla:441
    If vlatraceon() Then ' vla:442
        Call vlatracestep(810, vla_step_text(810)) ' vla:442
    End If
    gc_auto = range("b1") ' vla:444 src:1363
    vla_step = 811 ' vla:445
    If vlatraceon() Then ' vla:446
        Call vlatracestep(811, vla_step_text(811)) ' vla:446
    End If
    range("c5") = gc_auto ' vla:448 src:1364
    Exit Sub ' vla:449
vla_fail: ' vla:450
    Call vla_report_error ' vla:451
End Sub

Public Sub check_figures()
    Dim gfig_values As Variant ' vla:454
    Dim gfig_middle As Variant ' vla:455
    Dim gfig_spread As Variant ' vla:456
    On Error GoTo vla_fail ' vla:457
    vla_step = 813 ' vla:458
    If vlatraceon() Then ' vla:459
        Call vlatracestep(813, vla_step_text(813)) ' vla:459
    End If
    ' ---- instructions.txt:1368 To check-figures: ----
    Call vlaensuresheet("gfigures") ' vla:462 src:1380
    Call worksheets("gfigures").activate
    vla_step = 814 ' vla:463
    If vlatraceon() Then ' vla:464
        Call vlatracestep(814, vla_step_text(814)) ' vla:464
    End If
    Call range("a1:j10").clear
    vla_step = 815 ' vla:467
    If vlatraceon() Then ' vla:468
        Call vlatracestep(815, vla_step_text(815)) ' vla:468
    End If
    range("b2") = 2 ' vla:470 src:1382
    vla_step = 816 ' vla:471
    If vlatraceon() Then ' vla:472
        Call vlatracestep(816, vla_step_text(816)) ' vla:472
    End If
    range("b3") = 4 ' vla:474 src:1383
    vla_step = 817 ' vla:475
    If vlatraceon() Then ' vla:476
        Call vlatracestep(817, vla_step_text(817)) ' vla:476
    End If
    range("b4") = 6 ' vla:478 src:1384
    vla_step = 818 ' vla:479
    If vlatraceon() Then ' vla:480
        Call vlatracestep(818, vla_step_text(818)) ' vla:480
    End If
    range("b5") = 6 ' vla:482 src:1385
    vla_step = 819 ' vla:483
    If vlatraceon() Then ' vla:484
        Call vlatracestep(819, vla_step_text(819)) ' vla:484
    End If
    range("b6") = 7 ' vla:486 src:1386
    vla_step = 820 ' vla:487
    If vlatraceon() Then ' vla:488
        Call vlatracestep(820, vla_step_text(820)) ' vla:488
    End If
    Set gfig_values = range("b2:b6") ' vla:490 src:1387
    vla_step = 821 ' vla:491
    If vlatraceon() Then ' vla:492
        Call vlatracestep(821, vla_step_text(821)) ' vla:492
    End If
    gfig_middle = application.worksheetfunction.median(gfig_values) ' vla:494 src:1388
    vla_step = 822 ' vla:495
    If vlatraceon() Then ' vla:496
        Call vlatracestep(822, vla_step_text(822)) ' vla:496
    End If
    range("h1") = gfig_middle ' vla:498 src:1389
    vla_step = 823 ' vla:499
    If vlatraceon() Then ' vla:500
        Call vlatracestep(823, vla_step_text(823)) ' vla:500
    End If
    range("h2") = application.worksheetfunction.median(gfig_values) ' vla:502 src:1390
    vla_step = 824 ' vla:503
    If vlatraceon() Then ' vla:504
        Call vlatracestep(824, vla_step_text(824)) ' vla:504
    End If
    gfig_spread = application.worksheetfunction.stdev_s(gfig_values) ' vla:506 src:1391
    vla_step = 825 ' vla:507
    If vlatraceon() Then ' vla:508
        Call vlatracestep(825, vla_step_text(825)) ' vla:508
    End If
    range("h3") = gfig_spread ' vla:510 src:1392
    vla_step = 826 ' vla:511
    If vlatraceon() Then ' vla:512
        Call vlatracestep(826, vla_step_text(826)) ' vla:512
    End If
    range("h4") = application.worksheetfunction.stdev_p(gfig_values) ' vla:514 src:1393
    vla_step = 827 ' vla:515
    If vlatraceon() Then ' vla:516
        Call vlatracestep(827, vla_step_text(827)) ' vla:516
    End If
    range("d2") = "West" ' vla:518 src:1394
    vla_step = 828 ' vla:519
    If vlatraceon() Then ' vla:520
        Call vlatracestep(828, vla_step_text(828)) ' vla:520
    End If
    range("d3") = "East" ' vla:522 src:1395
    vla_step = 829 ' vla:523
    If vlatraceon() Then ' vla:524
        Call vlatracestep(829, vla_step_text(829)) ' vla:524
    End If
    range("d4") = "West" ' vla:526 src:1396
    vla_step = 830 ' vla:527
    If vlatraceon() Then ' vla:528
        Call vlatracestep(830, vla_step_text(830)) ' vla:528
    End If
    range("d5") = "West" ' vla:530 src:1397
    vla_step = 831 ' vla:531
    If vlatraceon() Then ' vla:532
        Call vlatracestep(831, vla_step_text(831)) ' vla:532
    End If
    range("e2") = 10 ' vla:534 src:1398
    vla_step = 832 ' vla:535
    If vlatraceon() Then ' vla:536
        Call vlatracestep(832, vla_step_text(832)) ' vla:536
    End If
    range("e3") = 20 ' vla:538 src:1399
    vla_step = 833 ' vla:539
    If vlatraceon() Then ' vla:540
        Call vlatracestep(833, vla_step_text(833)) ' vla:540
    End If
    range("e5") = 30 ' vla:542 src:1400
    vla_step = 834 ' vla:543
    If vlatraceon() Then ' vla:544
        Call vlatracestep(834, vla_step_text(834)) ' vla:544
    End If
    range("h5") = application.worksheetfunction.sumif(range("d2:d5"), "West", range("e2:e5")) ' vla:546 src:1401
    vla_step = 835 ' vla:547
    If vlatraceon() Then ' vla:548
        Call vlatracestep(835, vla_step_text(835)) ' vla:548
    End If
    range("h6") = application.worksheetfunction.averageif(range("d2:d5"), "West", range("e2:e5")) ' vla:550 src:1402
    vla_step = 836 ' vla:551
    If vlatraceon() Then ' vla:552
        Call vlatracestep(836, vla_step_text(836)) ' vla:552
    End If
    range("c2").numberformat = vlanumberformatcode("dollars", 2)
    vla_step = 837 ' vla:555
    If vlatraceon() Then ' vla:556
        Call vlatracestep(837, vla_step_text(837)) ' vla:556
    End If
    range("c2").Formula2 = "=B2*10"
    If (cells(rows.count, "b").end(xlup).row > range("c2").row) Then
        Call range("c2").resize(rowsize:=((cells(rows.count, "b").end(xlup).row - range("c2").row) + 1)).filldown
    End If
    vla_step = 838 ' vla:559
    If vlatraceon() Then ' vla:560
        Call vlatracestep(838, vla_step_text(838)) ' vla:560
    End If
    range("f1") = "head" ' vla:562 src:1405
    vla_step = 839 ' vla:563
    If vlatraceon() Then ' vla:564
        Call vlatracestep(839, vla_step_text(839)) ' vla:564
    End If
    range("g2") = 1 ' vla:566 src:1406
    vla_step = 840 ' vla:567
    If vlatraceon() Then ' vla:568
        Call vlatracestep(840, vla_step_text(840)) ' vla:568
    End If
    range("f2").Formula2 = "=B2*10"
    If (cells(rows.count, "g").end(xlup).row > range("f2").row) Then
        Call range("f2").resize(rowsize:=((cells(rows.count, "g").end(xlup).row - range("f2").row) + 1)).filldown
    End If
    vla_step = 841 ' vla:571
    If vlatraceon() Then ' vla:572
        Call vlatracestep(841, vla_step_text(841)) ' vla:572
    End If
    range("j2").Formula2 = "=B2*10"
    If (cells(rows.count, "i").end(xlup).row > range("j2").row) Then
        Call range("j2").resize(rowsize:=((cells(rows.count, "i").end(xlup).row - range("j2").row) + 1)).filldown
    End If
    Exit Sub ' vla:575
vla_fail: ' vla:576
    Call vla_report_error ' vla:577
End Sub

Public Sub main()
    Dim counter As Variant ' vla:580
    Dim grand As Variant ' vla:581
    Dim results As Variant ' vla:582
    Dim r As Variant ' vla:583
    Dim biggest As Variant ' vla:584
    Dim probe As Variant ' vla:585
    Dim round_check As Variant ' vla:586
    Dim thousand_check As Variant ' vla:587
    Dim f_last As Variant ' vla:588
    Dim echo_row As Variant ' vla:589
    Dim check_row As Variant ' vla:590
    Dim stripe_row As Variant ' vla:591
    Dim back_row As Variant ' vla:592
    Dim search_row As Variant ' vla:593
    Dim f As Variant ' vla:594
    Dim list_count As Variant ' vla:595
    Dim region As Variant ' vla:596
    Dim fuel As Variant ' vla:597
    Dim fee As Variant ' vla:598
    Dim full_commission As Variant ' vla:599
    Dim default_commission As Variant ' vla:600
    Dim pick_check As Variant ' vla:601
    Dim ax_price As Variant ' vla:602
    Dim key_count As Variant ' vla:603
    Dim price_sum As Variant ' vla:604
    Dim k As Variant ' vla:605
    Dim pair As Variant ' vla:606
    Dim qty_values As Variant ' vla:607
    Dim q As Variant ' vla:608
    Dim gtext_row As Variant ' vla:609
    Dim gtext_code As Variant ' vla:610
    Dim gtext_part As Variant ' vla:611
    Dim gtext_list As Variant ' vla:612
    Dim gtext_found As Variant ' vla:613
    On Error GoTo vla_fail ' vla:614
    vla_step = 1 ' vla:615
    If vlatraceon() Then ' vla:616
        Call vlatracestep(1, vla_step_text(1)) ' vla:616
    End If
    ' ---- instructions.txt:10 Work on sheet Output. ----
    Call vlaensuresheet("output") ' vla:619 src:18
    Call worksheets("output").activate
    vla_step = 2 ' vla:620
    If vlatraceon() Then ' vla:621
        Call vlatracestep(2, vla_step_text(2)) ' vla:621
    End If
    ' ---- instructions.txt:20 Turn off screen updating. ----
    application.screenupdating = False
    vla_step = 3 ' vla:625
    If vlatraceon() Then ' vla:626
        Call vlatracestep(3, vla_step_text(3)) ' vla:626
    End If
    Dim total As Double ' vla:628 src:21
    vla_step = 4 ' vla:629
    If vlatraceon() Then ' vla:630
        Call vlatracestep(4, vla_step_text(4)) ' vla:630
    End If
    total = 0 ' vla:632 src:22
    vla_step = 5 ' vla:633
    If vlatraceon() Then ' vla:634
        Call vlatracestep(5, vla_step_text(5)) ' vla:634
    End If
    range("a1") = "Test Report" ' vla:636 src:23
    vla_step = 6 ' vla:637
    If vlatraceon() Then ' vla:638
        Call vlatracestep(6, vla_step_text(6)) ' vla:638
    End If
    range("a1").font.bold = True
    vla_step = 7 ' vla:641
    If vlatraceon() Then ' vla:642
        Call vlatracestep(7, vla_step_text(7)) ' vla:642
    End If
    range("a1").font.size = 14
    vla_step = 8 ' vla:645
    If vlatraceon() Then ' vla:646
        Call vlatracestep(8, vla_step_text(8)) ' vla:646
    End If
    range("d1") = date() ' vla:648 src:26
    vla_step = 9 ' vla:649
    If vlatraceon() Then ' vla:650
        Call vlatracestep(9, vla_step_text(9)) ' vla:650
    End If
    ' ---- instructions.txt:28 Repeat 5 times: ----
    For counter = 1 To 5
        vla_step = 10 ' vla:654 src:28
        If vlatraceon() Then ' vla:655 src:28
            Call vlatracestep(10, vla_step_text(10)) ' vla:655 src:28
        End If
        total = (total + counter)
        vla_step = 11 ' vla:658 src:28
        If vlatraceon() Then ' vla:659 src:28
            Call vlatracestep(11, vla_step_text(11)) ' vla:659 src:28
        End If
        If ((counter Mod 2) = 0) Then ' vla:661 src:30
            Debug.Print ("even step " & counter) ' vla:661 src:30
        End If
    Next counter
    vla_step = 12 ' vla:662
    If vlatraceon() Then ' vla:663
        Call vlatracestep(12, vla_step_text(12)) ' vla:663
    End If
    ' ---- instructions.txt:32 Put total into cell B2. ----
    range("b2") = total ' vla:666 src:32
    vla_step = 13 ' vla:667
    If vlatraceon() Then ' vla:668
        Call vlatracestep(13, vla_step_text(13)) ' vla:668
    End If
    range("b3").Formula2 = "=B2*2"
    vla_step = 14 ' vla:671
    If vlatraceon() Then ' vla:672
        Call vlatracestep(14, vla_step_text(14)) ' vla:672
    End If
    range("b4") = application.worksheetfunction.sum(range("B2:B3")) ' vla:674 src:37
    vla_step = 15 ' vla:675
    If vlatraceon() Then ' vla:676
        Call vlatracestep(15, vla_step_text(15)) ' vla:676
    End If
    grand = application.worksheetfunction.sum(range("b2:b3")) ' vla:678 src:38
    vla_step = 16 ' vla:679
    If vlatraceon() Then ' vla:680
        Call vlatracestep(16, vla_step_text(16)) ' vla:680
    End If
    Debug.Print ("grand is " & grand) ' vla:682 src:39
    vla_step = 17 ' vla:683
    If vlatraceon() Then ' vla:684
        Call vlatracestep(17, vla_step_text(17)) ' vla:684
    End If
    ' ---- instructions.txt:41 If grand is greater than 40: ----
    If (grand > 40) Then ' vla:687 src:43
        vla_step = 18 ' vla:689 src:43
        If vlatraceon() Then ' vla:690 src:43
            Call vlatracestep(18, vla_step_text(18)) ' vla:690 src:43
        End If
        range("c4") = "PASS" ' vla:692 src:44
        vla_step = 19 ' vla:693 src:43
        If vlatraceon() Then ' vla:694 src:43
            Call vlatracestep(19, vla_step_text(19)) ' vla:694 src:43
        End If
        For counter = 1 To 2
            vla_step = 20 ' vla:697 src:45
            If vlatraceon() Then ' vla:698 src:45
                Call vlatracestep(20, vla_step_text(20)) ' vla:698 src:45
            End If
            Debug.Print ("pass check " & counter) ' vla:700 src:46
        Next counter
        vla_step = 21 ' vla:701 src:43
        If vlatraceon() Then ' vla:702 src:43
            Call vlatracestep(21, vla_step_text(21)) ' vla:702 src:43
        End If
        range("c4").font.bold = True
    Else
        vla_step = 22 ' vla:706 src:43
        If vlatraceon() Then ' vla:707 src:43
            Call vlatracestep(22, vla_step_text(22)) ' vla:707 src:43
        End If
        ' ---- instructions.txt:50 Otherwise: ----
        range("c4") = "CHECK" ' vla:710 src:51
    End If
    vla_step = 23 ' vla:711
    If vlatraceon() Then ' vla:712
        Call vlatracestep(23, vla_step_text(23)) ' vla:712
    End If
    ' ---- instructions.txt:53 If grand is at least 45, make cell C4 yellow. ----
    If (grand >= 45) Then ' vla:715 src:53
        range("c4").interior.color = vbyellow
    End If
    vla_step = 24 ' vla:716
    If vlatraceon() Then ' vla:717
        Call vlatracestep(24, vla_step_text(24)) ' vla:717
    End If
    Dim label As String ' vla:719 src:54
    vla_step = 25 ' vla:720
    If vlatraceon() Then ' vla:721
        Call vlatracestep(25, vla_step_text(25)) ' vla:721
    End If
    label = ("Total: " & total) ' vla:723 src:55
    vla_step = 26 ' vla:724
    If vlatraceon() Then ' vla:725
        Call vlatracestep(26, vla_step_text(26)) ' vla:725
    End If
    range("a6") = label ' vla:727 src:56
    vla_step = 27 ' vla:728
    If vlatraceon() Then ' vla:729
        Call vlatracestep(27, vla_step_text(27)) ' vla:729
    End If
    range("a6").font.color = vlacolor(hot_pink)
    vla_step = 28 ' vla:732
    If vlatraceon() Then ' vla:733
        Call vlatracestep(28, vla_step_text(28)) ' vla:733
    End If
    Set results = range("b2:b4") ' vla:735 src:58
    vla_step = 29 ' vla:736
    If vlatraceon() Then ' vla:737
        Call vlatracestep(29, vla_step_text(29)) ' vla:737
    End If
    For Each r In results ' vla:739 src:59
        Debug.Print r ' vla:739 src:59
    Next r
    vla_step = 30 ' vla:740
    If vlatraceon() Then ' vla:741
        Call vlatracestep(30, vla_step_text(30)) ' vla:741
    End If
    biggest = application.worksheetfunction.max(results) ' vla:743 src:60
    vla_step = 31 ' vla:744
    If vlatraceon() Then ' vla:745
        Call vlatracestep(31, vla_step_text(31)) ' vla:745
    End If
    Debug.Print ((("largest result is " & biggest) & ", label length ") & len(label)) ' vla:747 src:61
    vla_step = 32 ' vla:748
    If vlatraceon() Then ' vla:749
        Call vlatracestep(32, vla_step_text(32)) ' vla:749
    End If
    ' ---- instructions.txt:63 Repeat 3 times: ----
    For counter = 1 To 3
        vla_step = 33 ' vla:753 src:64
        If vlatraceon() Then ' vla:754 src:64
            Call vlatracestep(33, vla_step_text(33)) ' vla:754 src:64
        End If
        cells(counter, "e") = (counter * 10) ' vla:756 src:65
    Next counter
    vla_step = 34 ' vla:757
    If vlatraceon() Then ' vla:758
        Call vlatracestep(34, vla_step_text(34)) ' vla:758
    End If
    ' ---- instructions.txt:67 Set probe to cell in column E row 2. ----
    probe = cells(2, "e") ' vla:761 src:67
    vla_step = 35 ' vla:762
    If vlatraceon() Then ' vla:763
        Call vlatracestep(35, vla_step_text(35)) ' vla:763
    End If
    Dim num_col_check As String ' vla:765 src:68
    vla_step = 36 ' vla:766
    If vlatraceon() Then ' vla:767
        Call vlatracestep(36, vla_step_text(36)) ' vla:767
    End If
    If (cells(2, 5) = 20) Then ' vla:769 src:69
        num_col_check = "yes" ' vla:769 src:69
    End If
    vla_step = 37 ' vla:770
    If vlatraceon() Then ' vla:771
        Call vlatracestep(37, vla_step_text(37)) ' vla:771
    End If
    Dim value_word_check As String ' vla:773 src:70
    vla_step = 38 ' vla:774
    If vlatraceon() Then ' vla:775
        Call vlatracestep(38, vla_step_text(38)) ' vla:775
    End If
    If (cells(3, 5) = 30) Then ' vla:777 src:71
        value_word_check = "ok" ' vla:777 src:71
    End If
    vla_step = 39 ' vla:778
    If vlatraceon() Then ' vla:779
        Call vlatracestep(39, vla_step_text(39)) ' vla:779
    End If
    range("output!h16") = "bang" ' vla:781 src:72
    vla_step = 40 ' vla:782
    If vlatraceon() Then ' vla:783
        Call vlatracestep(40, vla_step_text(40)) ' vla:783
    End If
    range("a1").horizontalalignment = xlcenter
    vla_step = 41 ' vla:786
    If vlatraceon() Then ' vla:787
        Call vlatracestep(41, vla_step_text(41)) ' vla:787
    End If
    columns("d").columnwidth = 24
    vla_step = 42 ' vla:790
    If vlatraceon() Then ' vla:791
        Call vlatracestep(42, vla_step_text(42)) ' vla:791
    End If
    round_check = round(3.14159, 2) ' vla:793 src:75
    vla_step = 43 ' vla:794
    If vlatraceon() Then ' vla:795
        Call vlatracestep(43, vla_step_text(43)) ' vla:795
    End If
    Debug.Print ("rounded is " & round_check) ' vla:797 src:76
    vla_step = 44 ' vla:798
    If vlatraceon() Then ' vla:799
        Call vlatracestep(44, vla_step_text(44)) ' vla:799
    End If
    thousand_check = (1000 + 500) ' vla:801 src:77
    vla_step = 45 ' vla:802
    If vlatraceon() Then ' vla:803
        Call vlatracestep(45, vla_step_text(45)) ' vla:803
    End If
    Debug.Print ("thousands read as " & thousand_check) ' vla:805 src:78
    vla_step = 46 ' vla:806
    If vlatraceon() Then ' vla:807
        Call vlatracestep(46, vla_step_text(46)) ' vla:807
    End If
    Debug.Print ("probe is " & probe) ' vla:809 src:79
    vla_step = 47 ' vla:810
    If vlatraceon() Then ' vla:811
        Call vlatracestep(47, vla_step_text(47)) ' vla:811
    End If
    ' ---- instructions.txt:81 Create a number called countdown. ----
    Dim countdown As Double ' vla:814 src:81
    vla_step = 48 ' vla:815
    If vlatraceon() Then ' vla:816
        Call vlatracestep(48, vla_step_text(48)) ' vla:816
    End If
    countdown = 3 ' vla:818 src:82
    vla_step = 49 ' vla:819
    If vlatraceon() Then ' vla:820
        Call vlatracestep(49, vla_step_text(49)) ' vla:820
    End If
    Do While (countdown > 0) ' vla:822 src:83
        vla_step = 50 ' vla:823 src:83
        If vlatraceon() Then ' vla:824 src:83
            Call vlatracestep(50, vla_step_text(50)) ' vla:824 src:83
        End If
        Debug.Print ("countdown " & countdown) ' vla:826 src:84
        vla_step = 51 ' vla:827 src:83
        If vlatraceon() Then ' vla:828 src:83
            Call vlatracestep(51, vla_step_text(51)) ' vla:828 src:83
        End If
        countdown = (countdown - 1) ' vla:830 src:85
    Loop
    vla_step = 53 ' vla:831
    If vlatraceon() Then ' vla:832
        Call vlatracestep(53, vla_step_text(53)) ' vla:832
    End If
    ' ---- instructions.txt:92 Stamp. ----
    Call stamp ' vla:835 src:92
    vla_step = 54 ' vla:836
    If vlatraceon() Then ' vla:837
        Call vlatracestep(54, vla_step_text(54)) ' vla:837
    End If
    Call stamp(row_number:=2, value:="beta") ' vla:839 src:93
    vla_step = 55 ' vla:840
    If vlatraceon() Then ' vla:841
        Call vlatracestep(55, vla_step_text(55)) ' vla:841
    End If
    f_last = cells(rows.count, "f").end(xlup).row ' vla:843 src:94
    vla_step = 56 ' vla:844
    If vlatraceon() Then ' vla:845
        Call vlatracestep(56, vla_step_text(56)) ' vla:845
    End If
    Debug.Print ("column F filled to row " & f_last) ' vla:847 src:95
    vla_step = 57 ' vla:848
    If vlatraceon() Then ' vla:849
        Call vlatracestep(57, vla_step_text(57)) ' vla:849
    End If
    For echo_row = 1 To f_last ' vla:851 src:96
        Debug.Print ("echo " & echo_row) ' vla:851 src:96
    Next echo_row
    vla_step = 58 ' vla:852
    If vlatraceon() Then ' vla:853
        Call vlatracestep(58, vla_step_text(58)) ' vla:853
    End If
    For check_row = 1 To f_last ' vla:855 src:97
        vla_step = 59 ' vla:856 src:97
        If vlatraceon() Then ' vla:857 src:97
            Call vlatracestep(59, vla_step_text(59)) ' vla:857 src:97
        End If
        If (instr(1, cells(check_row, "f"), "ok", vbtextcompare) > 0) Then ' vla:859 src:98
            cells(check_row, "f").font.bold = True
        End If
    Next check_row
    vla_step = 60 ' vla:860
    If vlatraceon() Then ' vla:861
        Call vlatracestep(60, vla_step_text(60)) ' vla:861
    End If
    ' ---- instructions.txt:100 Count stripe-row from 1 to f-last step 2: ----
    For stripe_row = 1 To f_last Step 2 ' vla:864 src:103
        vla_step = 61 ' vla:865 src:103
        If vlatraceon() Then ' vla:866 src:103
            Call vlatracestep(61, vla_step_text(61)) ' vla:866 src:103
        End If
        cells(stripe_row, "f").font.bold = True
    Next stripe_row
    vla_step = 62 ' vla:869
    If vlatraceon() Then ' vla:870
        Call vlatracestep(62, vla_step_text(62)) ' vla:870
    End If
    ' ---- instructions.txt:106 Count back-row down from f-last to 1 step 2: ----
    For back_row = f_last To 1 Step (0 - 2) ' vla:873 src:106
        vla_step = 63 ' vla:874 src:106
        If vlatraceon() Then ' vla:875 src:106
            Call vlatracestep(63, vla_step_text(63)) ' vla:875 src:106
        End If
        Debug.Print ("back-row " & back_row) ' vla:877 src:107
    Next back_row
    vla_step = 64 ' vla:878
    If vlatraceon() Then ' vla:879
        Call vlatracestep(64, vla_step_text(64)) ' vla:879
    End If
    ' ---- instructions.txt:109 Count search-row from 1 to 10: ----
    For search_row = 1 To 10 ' vla:882 src:111
        vla_step = 65 ' vla:883 src:111
        If vlatraceon() Then ' vla:884 src:111
            Call vlatracestep(65, vla_step_text(65)) ' vla:884 src:111
        End If
        If (search_row = 3) Then ' vla:886 src:112
            Exit For ' vla:886 src:112
        End If
    Next search_row
    vla_step = 66 ' vla:887
    If vlatraceon() Then ' vla:888
        Call vlatracestep(66, vla_step_text(66)) ' vla:888
    End If
    ' ---- instructions.txt:114 Log "stopped at " joined with search-row. ----
    Debug.Print ("stopped at " & search_row) ' vla:891 src:114
    vla_step = 67 ' vla:892
    If vlatraceon() Then ' vla:893
        Call vlatracestep(67, vla_step_text(67)) ' vla:893
    End If
    ' ---- instructions.txt:116 Create a list called found-items. ----
    Dim found_items As Collection ' vla:896 src:119
    Set found_items = New Collection ' vla:896 src:119
    vla_step = 68 ' vla:897
    If vlatraceon() Then ' vla:898
        Call vlatracestep(68, vla_step_text(68)) ' vla:898
    End If
    For counter = 1 To 4
        vla_step = 69 ' vla:901 src:120
        If vlatraceon() Then ' vla:902 src:120
            Call vlatracestep(69, vla_step_text(69)) ' vla:902 src:120
        End If
        If (counter > 2) Then ' vla:904 src:121
            Call found_items.add((counter * 100)) ' vla:904 src:121
        End If
    Next counter
    vla_step = 70 ' vla:905
    If vlatraceon() Then ' vla:906
        Call vlatracestep(70, vla_step_text(70)) ' vla:906
    End If
    ' ---- instructions.txt:123 For each f in found-items, log "found " joined with f. ----
    For Each f In found_items ' vla:909 src:123
        Debug.Print ("found " & f) ' vla:909 src:123
    Next f
    vla_step = 71 ' vla:910
    If vlatraceon() Then ' vla:911
        Call vlatracestep(71, vla_step_text(71)) ' vla:911
    End If
    list_count = vlacount(found_items) ' vla:913 src:124
    vla_step = 72 ' vla:914
    If vlatraceon() Then ' vla:915
        Call vlatracestep(72, vla_step_text(72)) ' vla:915
    End If
    Debug.Print ("list holds " & list_count) ' vla:917 src:125
    vla_step = 73 ' vla:918
    If vlatraceon() Then ' vla:919
        Call vlatracestep(73, vla_step_text(73)) ' vla:919
    End If
    ' ---- instructions.txt:127 Create a text called verdict. ----
    Dim verdict As String ' vla:922 src:131
    vla_step = 74 ' vla:923
    If vlatraceon() Then ' vla:924
        Call vlatracestep(74, vla_step_text(74)) ' vla:924
    End If
    If (grand > 100) Then ' vla:926 src:132
        vla_step = 75 ' vla:928 src:132
        If vlatraceon() Then ' vla:929 src:132
            Call vlatracestep(75, vla_step_text(75)) ' vla:929 src:132
        End If
        verdict = "huge" ' vla:931 src:133
    ElseIf (grand > 40) Then
        vla_step = 76 ' vla:933 src:132
        If vlatraceon() Then ' vla:934 src:132
            Call vlatracestep(76, vla_step_text(76)) ' vla:934 src:132
        End If
        ' ---- instructions.txt:135 Otherwise, if grand is greater than 40: ----
        verdict = "solid" ' vla:937 src:136
    Else
        vla_step = 77 ' vla:939 src:132
        If vlatraceon() Then ' vla:940 src:132
            Call vlatracestep(77, vla_step_text(77)) ' vla:940 src:132
        End If
        ' ---- instructions.txt:138 Otherwise: ----
        verdict = "small" ' vla:943 src:139
    End If
    vla_step = 78 ' vla:944
    If vlatraceon() Then ' vla:945
        Call vlatracestep(78, vla_step_text(78)) ' vla:945
    End If
    ' ---- instructions.txt:141 Create a text called region-label. ----
    Dim region_label As String ' vla:948 src:141
    vla_step = 79 ' vla:949
    If vlatraceon() Then ' vla:950
        Call vlatracestep(79, vla_step_text(79)) ' vla:950
    End If
    region = "South" ' vla:952 src:142
    vla_step = 80 ' vla:953
    If vlatraceon() Then ' vla:954
        Call vlatracestep(80, vla_step_text(80)) ' vla:954
    End If
    Select Case region ' vla:956 src:143
        Case "North"
            vla_step = 81 ' vla:958 src:143
            If vlatraceon() Then ' vla:959 src:143
                Call vlatracestep(81, vla_step_text(81)) ' vla:959 src:143
            End If
            region_label = "cold" ' vla:961 src:144
        Case "South", "East"
            vla_step = 82 ' vla:963 src:143
            If vlatraceon() Then ' vla:964 src:143
                Call vlatracestep(82, vla_step_text(82)) ' vla:964 src:143
            End If
            ' ---- instructions.txt:146 When it is "South" or "East": ----
            region_label = "warm" ' vla:967 src:147
        Case Else
            vla_step = 83 ' vla:969 src:143
            If vlatraceon() Then ' vla:970 src:143
                Call vlatracestep(83, vla_step_text(83)) ' vla:970 src:143
            End If
            ' ---- instructions.txt:149 Otherwise: ----
            region_label = "unknown" ' vla:973 src:150
    End Select
    vla_step = 84 ' vla:974
    If vlatraceon() Then ' vla:975
        Call vlatracestep(84, vla_step_text(84)) ' vla:975
    End If
    ' ---- instructions.txt:152 Log "verdict " joined with verdict joined with ", region " joined w... ----
    Debug.Print ((("verdict " & verdict) & ", region ") & region_label) ' vla:978 src:152
    vla_step = 85 ' vla:979
    If vlatraceon() Then ' vla:980
        Call vlatracestep(85, vla_step_text(85)) ' vla:980
    End If
    ' ---- instructions.txt:154 Create a number called until-count. ----
    Dim until_count As Double ' vla:983 src:154
    vla_step = 86 ' vla:984
    If vlatraceon() Then ' vla:985
        Call vlatracestep(86, vla_step_text(86)) ' vla:985
    End If
    fuel = 3 ' vla:987 src:155
    vla_step = 87 ' vla:988
    If vlatraceon() Then ' vla:989
        Call vlatracestep(87, vla_step_text(87)) ' vla:989
    End If
    Do Until (fuel = 0) ' vla:991 src:156
        vla_step = 88 ' vla:992 src:156
        If vlatraceon() Then ' vla:993 src:156
            Call vlatracestep(88, vla_step_text(88)) ' vla:993 src:156
        End If
        fuel = (fuel - 1) ' vla:995 src:157
        vla_step = 89 ' vla:996 src:156
        If vlatraceon() Then ' vla:997 src:156
            Call vlatracestep(89, vla_step_text(89)) ' vla:997 src:156
        End If
        until_count = (until_count + 1)
    Loop
    vla_step = 90 ' vla:1000
    If vlatraceon() Then ' vla:1001
        Call vlatracestep(90, vla_step_text(90)) ' vla:1001
    End If
    ' ---- instructions.txt:160 Log "repeat-until ran " joined with until-count joined with " times". ----
    Debug.Print (("repeat-until ran " & until_count) & " times") ' vla:1004 src:160
    vla_step = 91 ' vla:1005
    If vlatraceon() Then ' vla:1006
        Call vlatracestep(91, vla_step_text(91)) ' vla:1006
    End If
    ' ---- instructions.txt:162 Create a text called rescue. ----
    Dim rescue As String ' vla:1009 src:166
    vla_step = 92 ' vla:1010
    If vlatraceon() Then ' vla:1011
        Call vlatracestep(92, vla_step_text(92)) ' vla:1011
    End If
    On Error GoTo vla_tryf_1 ' vla:1013 src:167
    vla_step = 93 ' vla:1014 src:167
    If vlatraceon() Then ' vla:1015 src:167
        Call vlatracestep(93, vla_step_text(93)) ' vla:1015 src:167
    End If
    Call worksheets("nowhere-land").activate
    vla_step = 94 ' vla:1018 src:167
    If vlatraceon() Then ' vla:1019 src:167
        Call vlatracestep(94, vla_step_text(94)) ' vla:1019 src:167
    End If
    rescue = "unreachable" ' vla:1021 src:169
    GoTo vla_tryd_1 ' vla:1022 src:167
vla_tryf_1: ' vla:1023 src:167
    vla_problem = err.description ' vla:1024 src:167
    Resume vla_tryr_1 ' vla:1025 src:167
vla_tryr_1: ' vla:1026 src:167
    On Error GoTo vla_fail ' vla:1027 src:167
    vla_step = 95 ' vla:1028 src:167
    If vlatraceon() Then ' vla:1029 src:167
        Call vlatracestep(95, vla_step_text(95)) ' vla:1029 src:167
    End If
    ' ---- instructions.txt:171 If that fails: ----
    Debug.Print ("the problem was " & vla_problem) ' vla:1032 src:172
    vla_step = 96 ' vla:1033 src:167
    If vlatraceon() Then ' vla:1034 src:167
        Call vlatracestep(96, vla_step_text(96)) ' vla:1034 src:167
    End If
    If (Not (len(trim((vla_problem & ""))) = 0)) Then ' vla:1036 src:173
        rescue = "rescued" ' vla:1036 src:173
    End If
vla_tryd_1: ' vla:1037 src:167
    On Error GoTo vla_fail ' vla:1038 src:167
    vla_step = 97 ' vla:1039
    If vlatraceon() Then ' vla:1040
        Call vlatracestep(97, vla_step_text(97)) ' vla:1040
    End If
    ' ---- instructions.txt:175 Create a number called risk-free. ----
    Dim risk_free As Double ' vla:1043 src:175
    vla_step = 98 ' vla:1044
    If vlatraceon() Then ' vla:1045
        Call vlatracestep(98, vla_step_text(98)) ' vla:1045
    End If
    On Error GoTo vla_tryf_2 ' vla:1047 src:176
    vla_step = 99 ' vla:1048 src:176
    If vlatraceon() Then ' vla:1049 src:176
        Call vlatracestep(99, vla_step_text(99)) ' vla:1049 src:176
    End If
    risk_free = 7 ' vla:1051 src:177
    GoTo vla_tryd_2 ' vla:1052 src:176
vla_tryf_2: ' vla:1053 src:176
    vla_problem = err.description ' vla:1054 src:176
    Resume vla_tryr_2 ' vla:1055 src:176
vla_tryr_2: ' vla:1056 src:176
    On Error GoTo vla_fail ' vla:1057 src:176
    vla_step = 100 ' vla:1058 src:176
    If vlatraceon() Then ' vla:1059 src:176
        Call vlatracestep(100, vla_step_text(100)) ' vla:1059 src:176
    End If
    ' ---- instructions.txt:179 If that fails: ----
    risk_free = -1 ' vla:1062 src:180
vla_tryd_2: ' vla:1063 src:176
    On Error GoTo vla_fail ' vla:1064 src:176
    vla_step = 101 ' vla:1065
    If vlatraceon() Then ' vla:1066
        Call vlatracestep(101, vla_step_text(101)) ' vla:1066
    End If
    ' ---- instructions.txt:182 Log "rescue " joined with rescue joined with ", risk-free " joined ... ----
    Debug.Print ((("rescue " & rescue) & ", risk-free ") & risk_free) ' vla:1069 src:182
    vla_step = 103 ' vla:1070
    If vlatraceon() Then ' vla:1071
        Call vlatracestep(103, vla_step_text(103)) ' vla:1071
    End If
    ' ---- instructions.txt:189 Set fee to tax of 100. ----
    fee = tax(100) ' vla:1074 src:189
    vla_step = 104 ' vla:1075
    If vlatraceon() Then ' vla:1076
        Call vlatracestep(104, vla_step_text(104)) ' vla:1076
    End If
    Debug.Print ("fee is " & fee) ' vla:1078 src:190
    vla_step = 105 ' vla:1079
    If vlatraceon() Then ' vla:1080
        Call vlatracestep(105, vla_step_text(105)) ' vla:1080
    End If
    Dim fee_size As String ' vla:1082 src:191
    vla_step = 106 ' vla:1083
    If vlatraceon() Then ' vla:1084
        Call vlatracestep(106, vla_step_text(106)) ' vla:1084
    End If
    If (tax(50) > 3) Then ' vla:1086 src:192
        vla_step = 107 ' vla:1088 src:192
        If vlatraceon() Then ' vla:1089 src:192
            Call vlatracestep(107, vla_step_text(107)) ' vla:1089 src:192
        End If
        fee_size = "big" ' vla:1091 src:193
    Else
        vla_step = 108 ' vla:1093 src:192
        If vlatraceon() Then ' vla:1094 src:192
            Call vlatracestep(108, vla_step_text(108)) ' vla:1094 src:192
        End If
        ' ---- instructions.txt:195 Otherwise: ----
        fee_size = "small" ' vla:1097 src:196
    End If
    vla_step = 113 ' vla:1098
    If vlatraceon() Then ' vla:1099
        Call vlatracestep(113, vla_step_text(113)) ' vla:1099
    End If
    ' ---- instructions.txt:208 Set full-commission to commission using sale of 2000 and rate of 10%. ----
    full_commission = commission(sale:=2000, rate:=(10 / 100)) ' vla:1102 src:208
    vla_step = 114 ' vla:1103
    If vlatraceon() Then ' vla:1104
        Call vlatracestep(114, vla_step_text(114)) ' vla:1104
    End If
    default_commission = commission(sale:=600) ' vla:1106 src:209
    vla_step = 115 ' vla:1107
    If vlatraceon() Then ' vla:1108
        Call vlatracestep(115, vla_step_text(115)) ' vla:1108
    End If
    Debug.Print ((("commissions " & full_commission) & " / ") & default_commission) ' vla:1110 src:210
    vla_step = 116 ' vla:1111
    If vlatraceon() Then ' vla:1112
        Call vlatracestep(116, vla_step_text(116)) ' vla:1112
    End If
    ' ---- instructions.txt:212 Try: ----
    On Error GoTo vla_tryf_3 ' vla:1115 src:219
    vla_step = 117 ' vla:1116 src:219
    If vlatraceon() Then ' vla:1117 src:219
        Call vlatracestep(117, vla_step_text(117)) ' vla:1117 src:219
    End If
    Call vlachecksheetname("Q1 Data")
    Call vlachecksheetabsent("Q1 Data")
    Call worksheets.add
    activesheet.name = "Q1 Data"
    GoTo vla_tryd_3 ' vla:1120 src:219
vla_tryf_3: ' vla:1121 src:219
    vla_problem = err.description ' vla:1122 src:219
    Resume vla_tryr_3 ' vla:1123 src:219
vla_tryr_3: ' vla:1124 src:219
    On Error GoTo vla_fail ' vla:1125 src:219
vla_tryd_3: ' vla:1126 src:219
    On Error GoTo vla_fail ' vla:1127 src:219
    vla_step = 118 ' vla:1128
    If vlatraceon() Then ' vla:1129
        Call vlatracestep(118, vla_step_text(118)) ' vla:1129
    End If
    ' ---- instructions.txt:222 Go to sheet Output. ----
    Call worksheets("output").activate
    vla_step = 119 ' vla:1133
    If vlatraceon() Then ' vla:1134
        Call vlatracestep(119, vla_step_text(119)) ' vla:1134
    End If
    range("'Q1 Data'!A1") = "spaced" ' vla:1136 src:223
    vla_step = 120 ' vla:1137
    If vlatraceon() Then ' vla:1138
        Call vlatracestep(120, vla_step_text(120)) ' vla:1138
    End If
    Dim spaced_check As String ' vla:1140 src:224
    vla_step = 121 ' vla:1141
    If vlatraceon() Then ' vla:1142
        Call vlatracestep(121, vla_step_text(121)) ' vla:1142
    End If
    spaced_check = range("'Q1 Data'!A1") ' vla:1144 src:225
    vla_step = 123 ' vla:1145
    If vlatraceon() Then ' vla:1146
        Call vlatracestep(123, vla_step_text(123)) ' vla:1146
    End If
    ' ---- instructions.txt:232 Create a number called growth-check. ----
    Dim growth_check As Double ' vla:1149 src:237
    vla_step = 124 ' vla:1150
    If vlatraceon() Then ' vla:1151
        Call vlatracestep(124, vla_step_text(124)) ' vla:1151
    End If
    growth_check = 200 ' vla:1153 src:238
    vla_step = 125 ' vla:1154
    If vlatraceon() Then ' vla:1155
        Call vlatracestep(125, vla_step_text(125)) ' vla:1155
    End If
    growth_check = (growth_check * (1 + (10 / 100))) ' vla:1157 src:239
    vla_step = 126 ' vla:1158
    If vlatraceon() Then ' vla:1159
        Call vlatracestep(126, vla_step_text(126)) ' vla:1159
    End If
    growth_check = (growth_check * (1 + (50 / 100))) ' vla:1161 src:240
    vla_step = 127 ' vla:1162
    If vlatraceon() Then ' vla:1163
        Call vlatracestep(127, vla_step_text(127)) ' vla:1163
    End If
    growth_check = (growth_check * (1 + (100 / 100))) ' vla:1165 src:241
    vla_step = 128 ' vla:1166
    If vlatraceon() Then ' vla:1167
        Call vlatracestep(128, vla_step_text(128)) ' vla:1167
    End If
    growth_check = (growth_check * (1 - (75 / 100))) ' vla:1169 src:242
    vla_step = 129 ' vla:1170
    If vlatraceon() Then ' vla:1171
        Call vlatracestep(129, vla_step_text(129)) ' vla:1171
    End If
    pick_check = vlaitem(found_items, 2) ' vla:1173 src:243
    vla_step = 130 ' vla:1174
    If vlatraceon() Then ' vla:1175
        Call vlatracestep(130, vla_step_text(130)) ' vla:1175
    End If
    Debug.Print ((((("grew to " & growth_check) & ", picked ") & pick_check) & ", first ") & vlafirst(found_items)) ' vla:1177 src:244
    vla_step = 131 ' vla:1178
    If vlatraceon() Then ' vla:1179
        Call vlatracestep(131, vla_step_text(131)) ' vla:1179
    End If
    Dim quote_check As String ' vla:1181 src:245
    vla_step = 132 ' vla:1182
    If vlatraceon() Then ' vla:1183
        Call vlatracestep(132, vla_step_text(132)) ' vla:1183
    End If
    quote_check = "He said ""ok""" ' vla:1185 src:246
    vla_step = 133 ' vla:1186
    If vlatraceon() Then ' vla:1187
        Call vlatracestep(133, vla_step_text(133)) ' vla:1187
    End If
    ' ---- instructions.txt:248 Create a lookup called prices. ----
    Dim prices As Object ' vla:1190 src:260
    Set prices = vladictnew() ' vla:1190 src:260
    vla_step = 134 ' vla:1191
    If vlatraceon() Then ' vla:1192
        Call vlatracestep(134, vla_step_text(134)) ' vla:1192
    End If
    Call vladictset(prices, "ax-7", 100) ' vla:1194 src:261
    vla_step = 135 ' vla:1195
    If vlatraceon() Then ' vla:1196
        Call vlatracestep(135, vla_step_text(135)) ' vla:1196
    End If
    Call vladictset(prices, "bx-2", 250) ' vla:1198 src:262
    vla_step = 136 ' vla:1199
    If vlatraceon() Then ' vla:1200
        Call vlatracestep(136, vla_step_text(136)) ' vla:1200
    End If
    Call vladictset(prices, "AX-7", 120) ' vla:1202 src:263
    vla_step = 137 ' vla:1203
    If vlatraceon() Then ' vla:1204
        Call vlatracestep(137, vla_step_text(137)) ' vla:1204
    End If
    ax_price = vladictget(prices, "ax-7") ' vla:1206 src:264
    vla_step = 138 ' vla:1207
    If vlatraceon() Then ' vla:1208
        Call vlatracestep(138, vla_step_text(138)) ' vla:1208
    End If
    key_count = vlacount(vladictkeys(prices)) ' vla:1210 src:265
    vla_step = 139 ' vla:1211
    If vlatraceon() Then ' vla:1212
        Call vlatracestep(139, vla_step_text(139)) ' vla:1212
    End If
    price_sum = (vladictget(prices, "ax-7") + vladictget(prices, "bx-2")) ' vla:1214 src:266
    vla_step = 140 ' vla:1215
    If vlatraceon() Then ' vla:1216
        Call vlatracestep(140, vla_step_text(140)) ' vla:1216
    End If
    Dim key_list As String ' vla:1218 src:267
    vla_step = 141 ' vla:1219
    If vlatraceon() Then ' vla:1220
        Call vlatracestep(141, vla_step_text(141)) ' vla:1220
    End If
    For Each k In vladictkeys(prices) ' vla:1222 src:268
        vla_step = 142 ' vla:1223 src:268
        If vlatraceon() Then ' vla:1224 src:268
            Call vlatracestep(142, vla_step_text(142)) ' vla:1224 src:268
        End If
        key_list = (key_list & k) ' vla:1226 src:269
    Next k
    vla_step = 143 ' vla:1227
    If vlatraceon() Then ' vla:1228
        Call vlatracestep(143, vla_step_text(143)) ' vla:1228
    End If
    ' ---- instructions.txt:271 Create a text called price-verdict. ----
    Dim price_verdict As String ' vla:1231 src:271
    vla_step = 144 ' vla:1232
    If vlatraceon() Then ' vla:1233
        Call vlatracestep(144, vla_step_text(144)) ' vla:1233
    End If
    If (vladictget(prices, "bx-2") > 200) Then ' vla:1235 src:272
        price_verdict = "steep" ' vla:1235 src:272
    End If
    vla_step = 145 ' vla:1236
    If vlatraceon() Then ' vla:1237
        Call vlatracestep(145, vla_step_text(145)) ' vla:1237
    End If
    Dim pair_trace As String ' vla:1239 src:273
    vla_step = 146 ' vla:1240
    If vlatraceon() Then ' vla:1241
        Call vlatracestep(146, vla_step_text(146)) ' vla:1241
    End If
    For Each pair In vladictpairs(prices) ' vla:1243 src:274
        vla_step = 147 ' vla:1244 src:274
        If vlatraceon() Then ' vla:1245 src:274
            Call vlatracestep(147, vla_step_text(147)) ' vla:1245 src:274
        End If
        If (vlapairvalue(pair) > 200) Then ' vla:1247 src:275
            pair_trace = (pair_trace & vlapairkey(pair)) ' vla:1247 src:275
        End If
    Next pair
    vla_step = 148 ' vla:1248
    If vlatraceon() Then ' vla:1249
        Call vlatracestep(148, vla_step_text(148)) ' vla:1249
    End If
    ' ---- instructions.txt:277 Log "lookup: ax " joined with ax-price joined with ", keys " joined... ----
    Debug.Print ((((("lookup: ax " & ax_price) & ", keys ") & key_list) & ", pairs ") & pair_trace) ' vla:1252 src:277
    vla_step = 149 ' vla:1253
    If vlatraceon() Then ' vla:1254
        Call vlatracestep(149, vla_step_text(149)) ' vla:1254
    End If
    ' ---- instructions.txt:279 Put "Item" into cell J1. ----
    range("j1") = "Item" ' vla:1257 src:284
    vla_step = 150 ' vla:1258
    If vlatraceon() Then ' vla:1259
        Call vlatracestep(150, vla_step_text(150)) ' vla:1259
    End If
    range("k1") = "Amount" ' vla:1261 src:285
    vla_step = 151 ' vla:1262
    If vlatraceon() Then ' vla:1263
        Call vlatracestep(151, vla_step_text(151)) ' vla:1263
    End If
    range("j2") = "Widget" ' vla:1265 src:286
    vla_step = 152 ' vla:1266
    If vlatraceon() Then ' vla:1267
        Call vlatracestep(152, vla_step_text(152)) ' vla:1267
    End If
    range("k2") = 10 ' vla:1269 src:287
    vla_step = 153 ' vla:1270
    If vlatraceon() Then ' vla:1271
        Call vlatracestep(153, vla_step_text(153)) ' vla:1271
    End If
    range("j3") = "Gadget" ' vla:1273 src:288
    vla_step = 154 ' vla:1274
    If vlatraceon() Then ' vla:1275
        Call vlatracestep(154, vla_step_text(154)) ' vla:1275
    End If
    range("k3") = 20 ' vla:1277 src:289
    vla_step = 155 ' vla:1278
    If vlatraceon() Then ' vla:1279
        Call vlatracestep(155, vla_step_text(155)) ' vla:1279
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j1:k3"), xllistobjecthasheaders:=xlyes).name = "salestable"
    vla_step = 156 ' vla:1282
    If vlatraceon() Then ' vla:1283
        Call vlatracestep(156, vla_step_text(156)) ' vla:1283
    End If
    activesheet.listobjects("salestable").tablestyle = "tablestylemedium9"
    vla_step = 157 ' vla:1286
    If vlatraceon() Then ' vla:1287
        Call vlatracestep(157, vla_step_text(157)) ' vla:1287
    End If
    activesheet.listobjects("salestable").showtotals = True
    vla_step = 158 ' vla:1290
    If vlatraceon() Then ' vla:1291
        Call vlatracestep(158, vla_step_text(158)) ' vla:1291
    End If
    ' ---- instructions.txt:294 Put "X" into cell J5. ----
    range("j5") = "X" ' vla:1294 src:296
    vla_step = 159 ' vla:1295
    If vlatraceon() Then ' vla:1296
        Call vlatracestep(159, vla_step_text(159)) ' vla:1296
    End If
    range("j6") = "Y" ' vla:1298 src:297
    vla_step = 160 ' vla:1299
    If vlatraceon() Then ' vla:1300
        Call vlatracestep(160, vla_step_text(160)) ' vla:1300
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j5:j6"), xllistobjecthasheaders:=xlyes).name = "quiettable"
    vla_step = 161 ' vla:1303
    If vlatraceon() Then ' vla:1304
        Call vlatracestep(161, vla_step_text(161)) ' vla:1304
    End If
    activesheet.listobjects("quiettable").showtotals = True
    vla_step = 162 ' vla:1307
    If vlatraceon() Then ' vla:1308
        Call vlatracestep(162, vla_step_text(162)) ' vla:1308
    End If
    activesheet.listobjects("quiettable").showtotals = False
    vla_step = 163 ' vla:1311
    If vlatraceon() Then ' vla:1312
        Call vlatracestep(163, vla_step_text(163)) ' vla:1312
    End If
    ' ---- instructions.txt:302 Put "A" into cell J8. ----
    range("j8") = "A" ' vla:1315 src:304
    vla_step = 164 ' vla:1316
    If vlatraceon() Then ' vla:1317
        Call vlatracestep(164, vla_step_text(164)) ' vla:1317
    End If
    range("j9") = "B" ' vla:1319 src:305
    vla_step = 165 ' vla:1320
    If vlatraceon() Then ' vla:1321
        Call vlatracestep(165, vla_step_text(165)) ' vla:1321
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j8:j9"), xllistobjecthasheaders:=xlyes).name = "temptable"
    vla_step = 166 ' vla:1324
    If vlatraceon() Then ' vla:1325
        Call vlatracestep(166, vla_step_text(166)) ' vla:1325
    End If
    Call activesheet.listobjects("temptable").unlist
    vla_step = 167 ' vla:1328
    If vlatraceon() Then ' vla:1329
        Call vlatracestep(167, vla_step_text(167)) ' vla:1329
    End If
    ' ---- instructions.txt:309 Put "Item" into cell J11. ----
    range("j11") = "Item" ' vla:1332 src:312
    vla_step = 168 ' vla:1333
    If vlatraceon() Then ' vla:1334
        Call vlatracestep(168, vla_step_text(168)) ' vla:1334
    End If
    range("k11") = "Qty" ' vla:1336 src:313
    vla_step = 169 ' vla:1337
    If vlatraceon() Then ' vla:1338
        Call vlatracestep(169, vla_step_text(169)) ' vla:1338
    End If
    range("j12") = "Bolt" ' vla:1340 src:314
    vla_step = 170 ' vla:1341
    If vlatraceon() Then ' vla:1342
        Call vlatracestep(170, vla_step_text(170)) ' vla:1342
    End If
    range("k12") = 5 ' vla:1344 src:315
    vla_step = 171 ' vla:1345
    If vlatraceon() Then ' vla:1346
        Call vlatracestep(171, vla_step_text(171)) ' vla:1346
    End If
    range("j13") = "Nut" ' vla:1348 src:316
    vla_step = 172 ' vla:1349
    If vlatraceon() Then ' vla:1350
        Call vlatracestep(172, vla_step_text(172)) ' vla:1350
    End If
    range("k13") = 8 ' vla:1352 src:317
    vla_step = 173 ' vla:1353
    If vlatraceon() Then ' vla:1354
        Call vlatracestep(173, vla_step_text(173)) ' vla:1354
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j11:k13"), xllistobjecthasheaders:=xlyes).name = "edittable"
    vla_step = 174 ' vla:1357
    If vlatraceon() Then ' vla:1358
        Call vlatracestep(174, vla_step_text(174)) ' vla:1358
    End If
    Call activesheet.listobjects("edittable").listrows.add
    vla_step = 175 ' vla:1361
    If vlatraceon() Then ' vla:1362
        Call vlatracestep(175, vla_step_text(175)) ' vla:1362
    End If
    Call activesheet.listobjects("edittable").listrows(1).delete
    vla_step = 176 ' vla:1365
    If vlatraceon() Then ' vla:1366
        Call vlatracestep(176, vla_step_text(176)) ' vla:1366
    End If
    Set qty_values = activesheet.listobjects("edittable").listcolumns("qty").databodyrange
    vla_step = 177 ' vla:1369
    If vlatraceon() Then ' vla:1370
        Call vlatracestep(177, vla_step_text(177)) ' vla:1370
    End If
    For Each q In qty_values ' vla:1372 src:322
        Debug.Print q ' vla:1372 src:322
    Next q
    vla_step = 178 ' vla:1373
    If vlatraceon() Then ' vla:1374
        Call vlatracestep(178, vla_step_text(178)) ' vla:1374
    End If
    ' ---- instructions.txt:324 Put "Region" into cell N1. ----
    range("n1") = "Region" ' vla:1377 src:332
    vla_step = 179 ' vla:1378
    If vlatraceon() Then ' vla:1379
        Call vlatracestep(179, vla_step_text(179)) ' vla:1379
    End If
    range("o1") = "Product" ' vla:1381 src:333
    vla_step = 180 ' vla:1382
    If vlatraceon() Then ' vla:1383
        Call vlatracestep(180, vla_step_text(180)) ' vla:1383
    End If
    range("p1") = "Segment" ' vla:1385 src:334
    vla_step = 181 ' vla:1386
    If vlatraceon() Then ' vla:1387
        Call vlatracestep(181, vla_step_text(181)) ' vla:1387
    End If
    range("q1") = "Channel" ' vla:1389 src:335
    vla_step = 182 ' vla:1390
    If vlatraceon() Then ' vla:1391
        Call vlatracestep(182, vla_step_text(182)) ' vla:1391
    End If
    range("r1") = "Units" ' vla:1393 src:336
    vla_step = 183 ' vla:1394
    If vlatraceon() Then ' vla:1395
        Call vlatracestep(183, vla_step_text(183)) ' vla:1395
    End If
    range("s1") = "Revenue" ' vla:1397 src:337
    vla_step = 184 ' vla:1398
    If vlatraceon() Then ' vla:1399
        Call vlatracestep(184, vla_step_text(184)) ' vla:1399
    End If
    range("n2") = "North" ' vla:1401 src:338
    vla_step = 185 ' vla:1402
    If vlatraceon() Then ' vla:1403
        Call vlatracestep(185, vla_step_text(185)) ' vla:1403
    End If
    range("o2") = "Widget" ' vla:1405 src:339
    vla_step = 186 ' vla:1406
    If vlatraceon() Then ' vla:1407
        Call vlatracestep(186, vla_step_text(186)) ' vla:1407
    End If
    range("p2") = "Retail" ' vla:1409 src:340
    vla_step = 187 ' vla:1410
    If vlatraceon() Then ' vla:1411
        Call vlatracestep(187, vla_step_text(187)) ' vla:1411
    End If
    range("q2") = "Online" ' vla:1413 src:341
    vla_step = 188 ' vla:1414
    If vlatraceon() Then ' vla:1415
        Call vlatracestep(188, vla_step_text(188)) ' vla:1415
    End If
    range("r2") = 10 ' vla:1417 src:342
    vla_step = 189 ' vla:1418
    If vlatraceon() Then ' vla:1419
        Call vlatracestep(189, vla_step_text(189)) ' vla:1419
    End If
    range("s2") = 500 ' vla:1421 src:343
    vla_step = 190 ' vla:1422
    If vlatraceon() Then ' vla:1423
        Call vlatracestep(190, vla_step_text(190)) ' vla:1423
    End If
    range("n3") = "North" ' vla:1425 src:344
    vla_step = 191 ' vla:1426
    If vlatraceon() Then ' vla:1427
        Call vlatracestep(191, vla_step_text(191)) ' vla:1427
    End If
    range("o3") = "Gadget" ' vla:1429 src:345
    vla_step = 192 ' vla:1430
    If vlatraceon() Then ' vla:1431
        Call vlatracestep(192, vla_step_text(192)) ' vla:1431
    End If
    range("p3") = "Wholesale" ' vla:1433 src:346
    vla_step = 193 ' vla:1434
    If vlatraceon() Then ' vla:1435
        Call vlatracestep(193, vla_step_text(193)) ' vla:1435
    End If
    range("q3") = "Store" ' vla:1437 src:347
    vla_step = 194 ' vla:1438
    If vlatraceon() Then ' vla:1439
        Call vlatracestep(194, vla_step_text(194)) ' vla:1439
    End If
    range("r3") = 5 ' vla:1441 src:348
    vla_step = 195 ' vla:1442
    If vlatraceon() Then ' vla:1443
        Call vlatracestep(195, vla_step_text(195)) ' vla:1443
    End If
    range("s3") = 200 ' vla:1445 src:349
    vla_step = 196 ' vla:1446
    If vlatraceon() Then ' vla:1447
        Call vlatracestep(196, vla_step_text(196)) ' vla:1447
    End If
    range("n4") = "South" ' vla:1449 src:350
    vla_step = 197 ' vla:1450
    If vlatraceon() Then ' vla:1451
        Call vlatracestep(197, vla_step_text(197)) ' vla:1451
    End If
    range("o4") = "Widget" ' vla:1453 src:351
    vla_step = 198 ' vla:1454
    If vlatraceon() Then ' vla:1455
        Call vlatracestep(198, vla_step_text(198)) ' vla:1455
    End If
    range("p4") = "Wholesale" ' vla:1457 src:352
    vla_step = 199 ' vla:1458
    If vlatraceon() Then ' vla:1459
        Call vlatracestep(199, vla_step_text(199)) ' vla:1459
    End If
    range("q4") = "Online" ' vla:1461 src:353
    vla_step = 200 ' vla:1462
    If vlatraceon() Then ' vla:1463
        Call vlatracestep(200, vla_step_text(200)) ' vla:1463
    End If
    range("r4") = 20 ' vla:1465 src:354
    vla_step = 201 ' vla:1466
    If vlatraceon() Then ' vla:1467
        Call vlatracestep(201, vla_step_text(201)) ' vla:1467
    End If
    range("s4") = 900 ' vla:1469 src:355
    vla_step = 202 ' vla:1470
    If vlatraceon() Then ' vla:1471
        Call vlatracestep(202, vla_step_text(202)) ' vla:1471
    End If
    range("n5") = "South" ' vla:1473 src:356
    vla_step = 203 ' vla:1474
    If vlatraceon() Then ' vla:1475
        Call vlatracestep(203, vla_step_text(203)) ' vla:1475
    End If
    range("o5") = "Gadget" ' vla:1477 src:357
    vla_step = 204 ' vla:1478
    If vlatraceon() Then ' vla:1479
        Call vlatracestep(204, vla_step_text(204)) ' vla:1479
    End If
    range("p5") = "Retail" ' vla:1481 src:358
    vla_step = 205 ' vla:1482
    If vlatraceon() Then ' vla:1483
        Call vlatracestep(205, vla_step_text(205)) ' vla:1483
    End If
    range("q5") = "Store" ' vla:1485 src:359
    vla_step = 206 ' vla:1486
    If vlatraceon() Then ' vla:1487
        Call vlatracestep(206, vla_step_text(206)) ' vla:1487
    End If
    range("r5") = 8 ' vla:1489 src:360
    vla_step = 207 ' vla:1490
    If vlatraceon() Then ' vla:1491
        Call vlatracestep(207, vla_step_text(207)) ' vla:1491
    End If
    range("s5") = 300 ' vla:1493 src:361
    vla_step = 208 ' vla:1494
    If vlatraceon() Then ' vla:1495
        Call vlatracestep(208, vla_step_text(208)) ' vla:1495
    End If
    range("n6") = "East" ' vla:1497 src:362
    vla_step = 209 ' vla:1498
    If vlatraceon() Then ' vla:1499
        Call vlatracestep(209, vla_step_text(209)) ' vla:1499
    End If
    range("o6") = "Widget" ' vla:1501 src:363
    vla_step = 210 ' vla:1502
    If vlatraceon() Then ' vla:1503
        Call vlatracestep(210, vla_step_text(210)) ' vla:1503
    End If
    range("p6") = "Retail" ' vla:1505 src:364
    vla_step = 211 ' vla:1506
    If vlatraceon() Then ' vla:1507
        Call vlatracestep(211, vla_step_text(211)) ' vla:1507
    End If
    range("q6") = "Online" ' vla:1509 src:365
    vla_step = 212 ' vla:1510
    If vlatraceon() Then ' vla:1511
        Call vlatracestep(212, vla_step_text(212)) ' vla:1511
    End If
    range("r6") = 12 ' vla:1513 src:366
    vla_step = 213 ' vla:1514
    If vlatraceon() Then ' vla:1515
        Call vlatracestep(213, vla_step_text(213)) ' vla:1515
    End If
    range("s6") = 600 ' vla:1517 src:367
    vla_step = 214 ' vla:1518
    If vlatraceon() Then ' vla:1519
        Call vlatracestep(214, vla_step_text(214)) ' vla:1519
    End If
    range("n7") = "East" ' vla:1521 src:368
    vla_step = 215 ' vla:1522
    If vlatraceon() Then ' vla:1523
        Call vlatracestep(215, vla_step_text(215)) ' vla:1523
    End If
    range("o7") = "Gadget" ' vla:1525 src:369
    vla_step = 216 ' vla:1526
    If vlatraceon() Then ' vla:1527
        Call vlatracestep(216, vla_step_text(216)) ' vla:1527
    End If
    range("p7") = "Wholesale" ' vla:1529 src:370
    vla_step = 217 ' vla:1530
    If vlatraceon() Then ' vla:1531
        Call vlatracestep(217, vla_step_text(217)) ' vla:1531
    End If
    range("q7") = "Store" ' vla:1533 src:371
    vla_step = 218 ' vla:1534
    If vlatraceon() Then ' vla:1535
        Call vlatracestep(218, vla_step_text(218)) ' vla:1535
    End If
    range("r7") = 6 ' vla:1537 src:372
    vla_step = 219 ' vla:1538
    If vlatraceon() Then ' vla:1539
        Call vlatracestep(219, vla_step_text(219)) ' vla:1539
    End If
    range("s7") = 250 ' vla:1541 src:373
    vla_step = 220 ' vla:1542
    If vlatraceon() Then ' vla:1543
        Call vlatracestep(220, vla_step_text(220)) ' vla:1543
    End If
    ' ---- instructions.txt:375 Make a pivot table from N1:S7 at U1 called SalesPivot. ----
    Call vlapivotcreate(range("n1:s7"), range("u1"), "salespivot")
    vla_step = 221 ' vla:1547
    If vlatraceon() Then ' vla:1548
        Call vlatracestep(221, vla_step_text(221)) ' vla:1548
    End If
    Call vlapivotsetorientation("salespivot", Array("region", "product"), "row")
    vla_step = 222 ' vla:1551
    If vlatraceon() Then ' vla:1552
        Call vlatracestep(222, vla_step_text(222)) ' vla:1552
    End If
    Call vlapivotsetorientation("salespivot", Array("segment"), "column")
    vla_step = 223 ' vla:1555
    If vlatraceon() Then ' vla:1556
        Call vlatracestep(223, vla_step_text(223)) ' vla:1556
    End If
    Call vlapivotsetorientation("salespivot", Array("channel"), "filter")
    vla_step = 224 ' vla:1559
    If vlatraceon() Then ' vla:1560
        Call vlatracestep(224, vla_step_text(224)) ' vla:1560
    End If
    ' ---- instructions.txt:386 Add Revenue to pivot SalesPivot as a sum. ----
    Call vlapivotaddvalues("salespivot", Array("revenue"), "sum")
    vla_step = 225 ' vla:1564
    If vlatraceon() Then ' vla:1565
        Call vlatracestep(225, vla_step_text(225)) ' vla:1565
    End If
    Call vlapivotaddvalues("salespivot", Array("revenue", "units"), "count")
    vla_step = 226 ' vla:1568
    If vlatraceon() Then ' vla:1569
        Call vlatracestep(226, vla_step_text(226)) ' vla:1569
    End If
    Call vlapivotaddvalues("salespivot", Array("units"), "average")
    vla_step = 227 ' vla:1572
    If vlatraceon() Then ' vla:1573
        Call vlatracestep(227, vla_step_text(227)) ' vla:1573
    End If
    ' ---- instructions.txt:396 Make a pivot table from N1:S7 at N20 called FullPivot with rows of ... ----
    Call vlapivotcreate(range("n1:s7"), range("n20"), "fullpivot")
    Call vlapivotsetorientation("fullpivot", Array("region", "product", "channel"), "row")
    Call vlapivotsetorientation("fullpivot", Array("segment"), "column")
    Call vlapivotaddvalues("fullpivot", Array("revenue", "units"), "sum")
    vla_step = 228 ' vla:1577
    If vlatraceon() Then ' vla:1578
        Call vlatracestep(228, vla_step_text(228)) ' vla:1578
    End If
    ' ---- instructions.txt:408 Refresh pivot SalesPivot. ----
    Call vlapivotrefresh("salespivot")
    vla_step = 229 ' vla:1582
    If vlatraceon() Then ' vla:1583
        Call vlatracestep(229, vla_step_text(229)) ' vla:1583
    End If
    Call vlapivotrefreshall
    vla_step = 230 ' vla:1586
    If vlatraceon() Then ' vla:1587
        Call vlatracestep(230, vla_step_text(230)) ' vla:1587
    End If
    Call vlapivotsetshowdetail("salespivot", Array("region"), False)
    vla_step = 231 ' vla:1590
    If vlatraceon() Then ' vla:1591
        Call vlatracestep(231, vla_step_text(231)) ' vla:1591
    End If
    ' ---- instructions.txt:419 Collapse Product in pivot FullPivot. ----
    Call vlapivotsetshowdetail("fullpivot", Array("product"), False)
    vla_step = 232 ' vla:1595
    If vlatraceon() Then ' vla:1596
        Call vlatracestep(232, vla_step_text(232)) ' vla:1596
    End If
    Call vlapivotsetshowdetail("fullpivot", Array("product"), True)
    vla_step = 233 ' vla:1599
    If vlatraceon() Then ' vla:1600
        Call vlatracestep(233, vla_step_text(233)) ' vla:1600
    End If
    ' ---- instructions.txt:430 Show pivot SalesPivot in tabular form. ----
    Call vlapivotsetrowlayout("salespivot", "tabular")
    vla_step = 234 ' vla:1604
    If vlatraceon() Then ' vla:1605
        Call vlatracestep(234, vla_step_text(234)) ' vla:1605
    End If
    Call vlapivotsetrowlayout("salespivot", "compact")
    vla_step = 235 ' vla:1608
    If vlatraceon() Then ' vla:1609
        Call vlatracestep(235, vla_step_text(235)) ' vla:1609
    End If
    Call vlapivotsetrowlayout("fullpivot", "outline")
    vla_step = 236 ' vla:1612
    If vlatraceon() Then ' vla:1613
        Call vlatracestep(236, vla_step_text(236)) ' vla:1613
    End If
    ' ---- instructions.txt:447 Hide subtotals for Region in pivot SalesPivot. ----
    Call vlapivotsetsubtotals("salespivot", Array("region"), False)
    vla_step = 237 ' vla:1617
    If vlatraceon() Then ' vla:1618
        Call vlatracestep(237, vla_step_text(237)) ' vla:1618
    End If
    Call vlapivotsetsubtotals("fullpivot", Array("product"), False)
    vla_step = 238 ' vla:1621
    If vlatraceon() Then ' vla:1622
        Call vlatracestep(238, vla_step_text(238)) ' vla:1622
    End If
    Call vlapivotsetsubtotals("fullpivot", Array("product"), True)
    vla_step = 239 ' vla:1625
    If vlatraceon() Then ' vla:1626
        Call vlatracestep(239, vla_step_text(239)) ' vla:1626
    End If
    ' ---- instructions.txt:457 Add a blank row after Region in pivot SalesPivot. ----
    Call vlapivotsetblankline("salespivot", Array("region"), True)
    vla_step = 240 ' vla:1630
    If vlatraceon() Then ' vla:1631
        Call vlatracestep(240, vla_step_text(240)) ' vla:1631
    End If
    Call vlapivotsetblankline("fullpivot", Array("product"), True)
    vla_step = 241 ' vla:1634
    If vlatraceon() Then ' vla:1635
        Call vlatracestep(241, vla_step_text(241)) ' vla:1635
    End If
    Call vlapivotsetblankline("fullpivot", Array("product"), False)
    vla_step = 242 ' vla:1638
    If vlatraceon() Then ' vla:1639
        Call vlatracestep(242, vla_step_text(242)) ' vla:1639
    End If
    ' ---- instructions.txt:461 Sort Region in pivot SalesPivot descending. ----
    Call vlapivotsort("salespivot", "region", "descending", "")
    vla_step = 243 ' vla:1643
    If vlatraceon() Then ' vla:1644
        Call vlatracestep(243, vla_step_text(243)) ' vla:1644
    End If
    Call vlapivotsort("salespivot", "region", "ascending", "")
    vla_step = 244 ' vla:1647
    If vlatraceon() Then ' vla:1648
        Call vlatracestep(244, vla_step_text(244)) ' vla:1648
    End If
    Call vlapivotsort("fullpivot", "product", "descending", "revenue")
    vla_step = 245 ' vla:1651
    If vlatraceon() Then ' vla:1652
        Call vlatracestep(245, vla_step_text(245)) ' vla:1652
    End If
    Call vlapivotsort("fullpivot", "product", "ascending", "revenue")
    vla_step = 246 ' vla:1655
    If vlatraceon() Then ' vla:1656
        Call vlatracestep(246, vla_step_text(246)) ' vla:1656
    End If
    ' ---- instructions.txt:481 Make a pivot table from N1:S7 at N200 called RenameMePivot. ----
    Call vlapivotcreate(range("n1:s7"), range("n200"), "renamemepivot")
    vla_step = 247 ' vla:1660
    If vlatraceon() Then ' vla:1661
        Call vlatracestep(247, vla_step_text(247)) ' vla:1661
    End If
    Call vlapivotrename("renamemepivot", "renamedpivot")
    vla_step = 248 ' vla:1664
    If vlatraceon() Then ' vla:1665
        Call vlatracestep(248, vla_step_text(248)) ' vla:1665
    End If
    ' ---- instructions.txt:493 Make a pivot table from N1:S7 at N220 called ClearMePivot. ----
    Call vlapivotcreate(range("n1:s7"), range("n220"), "clearmepivot")
    vla_step = 249 ' vla:1669
    If vlatraceon() Then ' vla:1670
        Call vlatracestep(249, vla_step_text(249)) ' vla:1670
    End If
    Call vlapivotsetorientation("clearmepivot", Array("region"), "row")
    vla_step = 250 ' vla:1673
    If vlatraceon() Then ' vla:1674
        Call vlatracestep(250, vla_step_text(250)) ' vla:1674
    End If
    Call vlapivotaddvalues("clearmepivot", Array("revenue"), "sum")
    vla_step = 251 ' vla:1677
    If vlatraceon() Then ' vla:1678
        Call vlatracestep(251, vla_step_text(251)) ' vla:1678
    End If
    Call vlapivotclear("clearmepivot")
    vla_step = 252 ' vla:1681
    If vlatraceon() Then ' vla:1682
        Call vlatracestep(252, vla_step_text(252)) ' vla:1682
    End If
    ' ---- instructions.txt:498 Make a pivot table from N1:S7 at N240 called RemoveFieldMePivot. ----
    Call vlapivotcreate(range("n1:s7"), range("n240"), "removefieldmepivot")
    vla_step = 253 ' vla:1686
    If vlatraceon() Then ' vla:1687
        Call vlatracestep(253, vla_step_text(253)) ' vla:1687
    End If
    Call vlapivotsetorientation("removefieldmepivot", Array("region", "product"), "row")
    vla_step = 254 ' vla:1690
    If vlatraceon() Then ' vla:1691
        Call vlatracestep(254, vla_step_text(254)) ' vla:1691
    End If
    Call vlapivotsetorientation("removefieldmepivot", Array("product"), "hidden")
    vla_step = 255 ' vla:1694
    If vlatraceon() Then ' vla:1695
        Call vlatracestep(255, vla_step_text(255)) ' vla:1695
    End If
    ' ---- instructions.txt:502 Put "West" into cell N8. ----
    range("n8") = "West" ' vla:1698 src:506
    vla_step = 256 ' vla:1699
    If vlatraceon() Then ' vla:1700
        Call vlatracestep(256, vla_step_text(256)) ' vla:1700
    End If
    range("o8") = "Widget" ' vla:1702 src:507
    vla_step = 257 ' vla:1703
    If vlatraceon() Then ' vla:1704
        Call vlatracestep(257, vla_step_text(257)) ' vla:1704
    End If
    range("p8") = "Retail" ' vla:1706 src:508
    vla_step = 258 ' vla:1707
    If vlatraceon() Then ' vla:1708
        Call vlatracestep(258, vla_step_text(258)) ' vla:1708
    End If
    range("q8") = "Online" ' vla:1710 src:509
    vla_step = 259 ' vla:1711
    If vlatraceon() Then ' vla:1712
        Call vlatracestep(259, vla_step_text(259)) ' vla:1712
    End If
    range("r8") = 15 ' vla:1714 src:510
    vla_step = 260 ' vla:1715
    If vlatraceon() Then ' vla:1716
        Call vlatracestep(260, vla_step_text(260)) ' vla:1716
    End If
    range("s8") = 700 ' vla:1718 src:511
    vla_step = 261 ' vla:1719
    If vlatraceon() Then ' vla:1720
        Call vlatracestep(261, vla_step_text(261)) ' vla:1720
    End If
    ' ---- instructions.txt:513 Make a pivot table from N1:S7 at N260 called SourceTestPivot. ----
    Call vlapivotcreate(range("n1:s7"), range("n260"), "sourcetestpivot")
    vla_step = 262 ' vla:1724
    If vlatraceon() Then ' vla:1725
        Call vlatracestep(262, vla_step_text(262)) ' vla:1725
    End If
    Call vlapivotsetorientation("sourcetestpivot", Array("region"), "row")
    vla_step = 263 ' vla:1728
    If vlatraceon() Then ' vla:1729
        Call vlatracestep(263, vla_step_text(263)) ' vla:1729
    End If
    Call vlapivotchangesource("sourcetestpivot", range("n1:s8"))
    vla_step = 264 ' vla:1732
    If vlatraceon() Then ' vla:1733
        Call vlatracestep(264, vla_step_text(264)) ' vla:1733
    End If
    ' ---- instructions.txt:517 Make a pivot table from N1:S7 at N40 called TempPivot. ----
    Call vlapivotcreate(range("n1:s7"), range("n40"), "temppivot")
    vla_step = 265 ' vla:1737
    If vlatraceon() Then ' vla:1738
        Call vlatracestep(265, vla_step_text(265)) ' vla:1738
    End If
    Call vlapivotdelete("temppivot")
    vla_step = 266 ' vla:1741
    If vlatraceon() Then ' vla:1742
        Call vlatracestep(266, vla_step_text(266)) ' vla:1742
    End If
    ' ---- instructions.txt:523 Put grand into cell H1. ----
    range("h1") = grand ' vla:1745 src:524
    vla_step = 267 ' vla:1746
    If vlatraceon() Then ' vla:1747
        Call vlatracestep(267, vla_step_text(267)) ' vla:1747
    End If
    range("h2") = biggest ' vla:1749 src:525
    vla_step = 268 ' vla:1750
    If vlatraceon() Then ' vla:1751
        Call vlatracestep(268, vla_step_text(268)) ' vla:1751
    End If
    range("h3") = round_check ' vla:1753 src:526
    vla_step = 269 ' vla:1754
    If vlatraceon() Then ' vla:1755
        Call vlatracestep(269, vla_step_text(269)) ' vla:1755
    End If
    range("h4") = thousand_check ' vla:1757 src:527
    vla_step = 270 ' vla:1758
    If vlatraceon() Then ' vla:1759
        Call vlatracestep(270, vla_step_text(270)) ' vla:1759
    End If
    range("h5") = search_row ' vla:1761 src:528
    vla_step = 271 ' vla:1762
    If vlatraceon() Then ' vla:1763
        Call vlatracestep(271, vla_step_text(271)) ' vla:1763
    End If
    range("h6") = f_last ' vla:1765 src:529
    vla_step = 272 ' vla:1766
    If vlatraceon() Then ' vla:1767
        Call vlatracestep(272, vla_step_text(272)) ' vla:1767
    End If
    range("h7") = list_count ' vla:1769 src:530
    vla_step = 273 ' vla:1770
    If vlatraceon() Then ' vla:1771
        Call vlatracestep(273, vla_step_text(273)) ' vla:1771
    End If
    range("h8") = verdict ' vla:1773 src:531
    vla_step = 274 ' vla:1774
    If vlatraceon() Then ' vla:1775
        Call vlatracestep(274, vla_step_text(274)) ' vla:1775
    End If
    range("h9") = region_label ' vla:1777 src:532
    vla_step = 275 ' vla:1778
    If vlatraceon() Then ' vla:1779
        Call vlatracestep(275, vla_step_text(275)) ' vla:1779
    End If
    range("h10") = until_count ' vla:1781 src:533
    vla_step = 276 ' vla:1782
    If vlatraceon() Then ' vla:1783
        Call vlatracestep(276, vla_step_text(276)) ' vla:1783
    End If
    range("h11") = rescue ' vla:1785 src:534
    vla_step = 277 ' vla:1786
    If vlatraceon() Then ' vla:1787
        Call vlatracestep(277, vla_step_text(277)) ' vla:1787
    End If
    range("h12") = risk_free ' vla:1789 src:535
    vla_step = 278 ' vla:1790
    If vlatraceon() Then ' vla:1791
        Call vlatracestep(278, vla_step_text(278)) ' vla:1791
    End If
    range("h13") = fee ' vla:1793 src:536
    vla_step = 279 ' vla:1794
    If vlatraceon() Then ' vla:1795
        Call vlatracestep(279, vla_step_text(279)) ' vla:1795
    End If
    range("h14") = fee_size ' vla:1797 src:537
    vla_step = 280 ' vla:1798
    If vlatraceon() Then ' vla:1799
        Call vlatracestep(280, vla_step_text(280)) ' vla:1799
    End If
    range("h15") = num_col_check ' vla:1801 src:538
    vla_step = 281 ' vla:1802
    If vlatraceon() Then ' vla:1803
        Call vlatracestep(281, vla_step_text(281)) ' vla:1803
    End If
    range("h17") = value_word_check ' vla:1805 src:539
    vla_step = 282 ' vla:1806
    If vlatraceon() Then ' vla:1807
        Call vlatracestep(282, vla_step_text(282)) ' vla:1807
    End If
    range("h18") = spaced_check ' vla:1809 src:540
    vla_step = 283 ' vla:1810
    If vlatraceon() Then ' vla:1811
        Call vlatracestep(283, vla_step_text(283)) ' vla:1811
    End If
    range("h19") = full_commission ' vla:1813 src:541
    vla_step = 284 ' vla:1814
    If vlatraceon() Then ' vla:1815
        Call vlatracestep(284, vla_step_text(284)) ' vla:1815
    End If
    range("h20") = default_commission ' vla:1817 src:542
    vla_step = 285 ' vla:1818
    If vlatraceon() Then ' vla:1819
        Call vlatracestep(285, vla_step_text(285)) ' vla:1819
    End If
    range("h21") = growth_check ' vla:1821 src:543
    vla_step = 286 ' vla:1822
    If vlatraceon() Then ' vla:1823
        Call vlatracestep(286, vla_step_text(286)) ' vla:1823
    End If
    range("h22") = pick_check ' vla:1825 src:544
    vla_step = 287 ' vla:1826
    If vlatraceon() Then ' vla:1827
        Call vlatracestep(287, vla_step_text(287)) ' vla:1827
    End If
    range("h23") = quote_check ' vla:1829 src:545
    vla_step = 288 ' vla:1830
    If vlatraceon() Then ' vla:1831
        Call vlatracestep(288, vla_step_text(288)) ' vla:1831
    End If
    range("h24") = vat_rate() ' vla:1833 src:546
    vla_step = 289 ' vla:1834
    If vlatraceon() Then ' vla:1835
        Call vlatracestep(289, vla_step_text(289)) ' vla:1835
    End If
    range("h26") = ax_price ' vla:1837 src:547
    vla_step = 290 ' vla:1838
    If vlatraceon() Then ' vla:1839
        Call vlatracestep(290, vla_step_text(290)) ' vla:1839
    End If
    range("h27") = key_count ' vla:1841 src:548
    vla_step = 291 ' vla:1842
    If vlatraceon() Then ' vla:1843
        Call vlatracestep(291, vla_step_text(291)) ' vla:1843
    End If
    range("h28") = key_list ' vla:1845 src:549
    vla_step = 292 ' vla:1846
    If vlatraceon() Then ' vla:1847
        Call vlatracestep(292, vla_step_text(292)) ' vla:1847
    End If
    range("h29") = price_sum ' vla:1849 src:550
    vla_step = 293 ' vla:1850
    If vlatraceon() Then ' vla:1851
        Call vlatracestep(293, vla_step_text(293)) ' vla:1851
    End If
    range("h30") = price_verdict ' vla:1853 src:551
    vla_step = 294 ' vla:1854
    If vlatraceon() Then ' vla:1855
        Call vlatracestep(294, vla_step_text(294)) ' vla:1855
    End If
    range("h31") = pair_trace ' vla:1857 src:552
    vla_step = 295 ' vla:1858
    If vlatraceon() Then ' vla:1859
        Call vlatracestep(295, vla_step_text(295)) ' vla:1859
    End If
    ' ---- instructions.txt:554 (set! (range "h25") "vla-row") ----
    range("h25") = "vla-row" ' vla:1862 src:559
    vla_step = 296 ' vla:1863
    If vlatraceon() Then ' vla:1864
        Call vlatracestep(296, vla_step_text(296)) ' vla:1864
    End If
    For counter = 1 To 2
        vla_step = 297 ' vla:1867 src:560
        If vlatraceon() Then ' vla:1868 src:560
            Call vlatracestep(297, vla_step_text(297)) ' vla:1868 src:560
        End If
        Debug.Print counter ' vla:1870 src:561
    Next counter
    vla_step = 298 ' vla:1871
    If vlatraceon() Then ' vla:1872
        Call vlatracestep(298, vla_step_text(298)) ' vla:1872
    End If
    ' ---- instructions.txt:563 Work on sheet Demo. ----
    Call vlaensuresheet("demo") ' vla:1875 src:576
    Call worksheets("demo").activate
    vla_step = 299 ' vla:1876
    If vlatraceon() Then ' vla:1877
        Call vlatracestep(299, vla_step_text(299)) ' vla:1877
    End If
    ' ---- instructions.txt:578 Make cell A1 italic. ----
    range("a1").font.italic = True
    vla_step = 300 ' vla:1881
    If vlatraceon() Then ' vla:1882
        Call vlatracestep(300, vla_step_text(300)) ' vla:1882
    End If
    range("a2").interior.color = vbred
    vla_step = 301 ' vla:1885
    If vlatraceon() Then ' vla:1886
        Call vlatracestep(301, vla_step_text(301)) ' vla:1886
    End If
    range("a3").interior.color = vbyellow
    vla_step = 302 ' vla:1889
    If vlatraceon() Then ' vla:1890
        Call vlatracestep(302, vla_step_text(302)) ' vla:1890
    End If
    range("a4").font.color = vlacolor(hot_pink)
    vla_step = 303 ' vla:1893
    If vlatraceon() Then ' vla:1894
        Call vlatracestep(303, vla_step_text(303)) ' vla:1894
    End If
    range("a5").interior.color = vlacolor("#FF69B4")
    vla_step = 304 ' vla:1897
    If vlatraceon() Then ' vla:1898
        Call vlatracestep(304, vla_step_text(304)) ' vla:1898
    End If
    range("a5").interior.colorindex = xlnone
    vla_step = 305 ' vla:1901
    If vlatraceon() Then ' vla:1902
        Call vlatracestep(305, vla_step_text(305)) ' vla:1902
    End If
    range("a1:e10").borders.linestyle = xlcontinuous
    vla_step = 306 ' vla:1905
    If vlatraceon() Then ' vla:1906
        Call vlatracestep(306, vla_step_text(306)) ' vla:1906
    End If
    range("b1").numberformat = "$#,##0.00"
    vla_step = 307 ' vla:1909
    If vlatraceon() Then ' vla:1910
        Call vlatracestep(307, vla_step_text(307)) ' vla:1910
    End If
    range("b2").numberformat = "0.0%"
    vla_step = 308 ' vla:1913
    If vlatraceon() Then ' vla:1914
        Call vlatracestep(308, vla_step_text(308)) ' vla:1914
    End If
    range("b3").numberformat = "mm/dd/yyyy"
    vla_step = 309 ' vla:1917
    If vlatraceon() Then ' vla:1918
        Call vlatracestep(309, vla_step_text(309)) ' vla:1918
    End If
    range("a1").horizontalalignment = xlcenter
    vla_step = 310 ' vla:1921
    If vlatraceon() Then ' vla:1922
        Call vlatracestep(310, vla_step_text(310)) ' vla:1922
    End If
    range("a1:e1").horizontalalignment = xlcenter
    vla_step = 311 ' vla:1925
    If vlatraceon() Then ' vla:1926
        Call vlatracestep(311, vla_step_text(311)) ' vla:1926
    End If
    range("a2:a5").horizontalalignment = xlleft
    vla_step = 312 ' vla:1929
    If vlatraceon() Then ' vla:1930
        Call vlatracestep(312, vla_step_text(312)) ' vla:1930
    End If
    range("a6").horizontalalignment = xlcenter
    vla_step = 313 ' vla:1933
    If vlatraceon() Then ' vla:1934
        Call vlatracestep(313, vla_step_text(313)) ' vla:1934
    End If
    range("c1:c5").wraptext = True
    vla_step = 314 ' vla:1937
    If vlatraceon() Then ' vla:1938
        Call vlatracestep(314, vla_step_text(314)) ' vla:1938
    End If
    range("c1:c5").wraptext = False
    vla_step = 315 ' vla:1941
    If vlatraceon() Then ' vla:1942
        Call vlatracestep(315, vla_step_text(315)) ' vla:1942
    End If
    Call range("d1:d3").merge
    vla_step = 316 ' vla:1945
    If vlatraceon() Then ' vla:1946
        Call vlatracestep(316, vla_step_text(316)) ' vla:1946
    End If
    Call range("d1:d3").unmerge
    vla_step = 317 ' vla:1949
    If vlatraceon() Then ' vla:1950
        Call vlatracestep(317, vla_step_text(317)) ' vla:1950
    End If
    ' ---- instructions.txt:598 Set height of row 1 to 30. ----
    rows(1).rowheight = 30
    vla_step = 318 ' vla:1954
    If vlatraceon() Then ' vla:1955
        Call vlatracestep(318, vla_step_text(318)) ' vla:1955
    End If
    columns("a").columnwidth = 20
    vla_step = 319 ' vla:1958
    If vlatraceon() Then ' vla:1959
        Call vlatracestep(319, vla_step_text(319)) ' vla:1959
    End If
    Call columns("b").insert
    vla_step = 320 ' vla:1962
    If vlatraceon() Then ' vla:1963
        Call vlatracestep(320, vla_step_text(320)) ' vla:1963
    End If
    columns("c").hidden = True
    vla_step = 321 ' vla:1966
    If vlatraceon() Then ' vla:1967
        Call vlatracestep(321, vla_step_text(321)) ' vla:1967
    End If
    columns("c").hidden = False
    vla_step = 322 ' vla:1970
    If vlatraceon() Then ' vla:1971
        Call vlatracestep(322, vla_step_text(322)) ' vla:1971
    End If
    Call columns("b").delete
    vla_step = 323 ' vla:1974
    If vlatraceon() Then ' vla:1975
        Call vlatracestep(323, vla_step_text(323)) ' vla:1975
    End If
    Call rows(5).insert
    vla_step = 324 ' vla:1978
    If vlatraceon() Then ' vla:1979
        Call vlatracestep(324, vla_step_text(324)) ' vla:1979
    End If
    Call rows(1).insert
    vla_step = 325 ' vla:1982
    If vlatraceon() Then ' vla:1983
        Call vlatracestep(325, vla_step_text(325)) ' vla:1983
    End If
    Call rows(2).delete
    vla_step = 326 ' vla:1986
    If vlatraceon() Then ' vla:1987
        Call vlatracestep(326, vla_step_text(326)) ' vla:1987
    End If
    Call rows(3).delete
    vla_step = 327 ' vla:1990
    If vlatraceon() Then ' vla:1991
        Call vlatracestep(327, vla_step_text(327)) ' vla:1991
    End If
    rows(10).hidden = True
    vla_step = 328 ' vla:1994
    If vlatraceon() Then ' vla:1995
        Call vlatracestep(328, vla_step_text(328)) ' vla:1995
    End If
    rows(10).hidden = False
    vla_step = 329 ' vla:1998
    If vlatraceon() Then ' vla:1999
        Call vlatracestep(329, vla_step_text(329)) ' vla:1999
    End If
    Call rows(2).select
    activewindow.freezepanes = True
    vla_step = 330 ' vla:2002
    If vlatraceon() Then ' vla:2003
        Call vlatracestep(330, vla_step_text(330)) ' vla:2003
    End If
    activewindow.freezepanes = False
    vla_step = 331 ' vla:2006
    If vlatraceon() Then ' vla:2007
        Call vlatracestep(331, vla_step_text(331)) ' vla:2007
    End If
    Call columns("a").autofit
    vla_step = 332 ' vla:2010
    If vlatraceon() Then ' vla:2011
        Call vlatracestep(332, vla_step_text(332)) ' vla:2011
    End If
    Call cells.entirecolumn.autofit
    vla_step = 333 ' vla:2014
    If vlatraceon() Then ' vla:2015
        Call vlatracestep(333, vla_step_text(333)) ' vla:2015
    End If
    ' ---- instructions.txt:618 Put "Region" into cell G1. ----
    range("g1") = "Region" ' vla:2018 src:620
    vla_step = 334 ' vla:2019
    If vlatraceon() Then ' vla:2020
        Call vlatracestep(334, vla_step_text(334)) ' vla:2020
    End If
    range("h1") = "Amount" ' vla:2022 src:621
    vla_step = 335 ' vla:2023
    If vlatraceon() Then ' vla:2024
        Call vlatracestep(335, vla_step_text(335)) ' vla:2024
    End If
    range("i1") = "Notes" ' vla:2026 src:622
    vla_step = 336 ' vla:2027
    If vlatraceon() Then ' vla:2028
        Call vlatracestep(336, vla_step_text(336)) ' vla:2028
    End If
    range("g2") = "West" ' vla:2030 src:623
    vla_step = 337 ' vla:2031
    If vlatraceon() Then ' vla:2032
        Call vlatracestep(337, vla_step_text(337)) ' vla:2032
    End If
    range("h2") = 100 ' vla:2034 src:624
    vla_step = 338 ' vla:2035
    If vlatraceon() Then ' vla:2036
        Call vlatracestep(338, vla_step_text(338)) ' vla:2036
    End If
    range("i2") = "ok" ' vla:2038 src:625
    vla_step = 339 ' vla:2039
    If vlatraceon() Then ' vla:2040
        Call vlatracestep(339, vla_step_text(339)) ' vla:2040
    End If
    range("g3") = "East" ' vla:2042 src:626
    vla_step = 340 ' vla:2043
    If vlatraceon() Then ' vla:2044
        Call vlatracestep(340, vla_step_text(340)) ' vla:2044
    End If
    range("h3") = 250 ' vla:2046 src:627
    vla_step = 341 ' vla:2047
    If vlatraceon() Then ' vla:2048
        Call vlatracestep(341, vla_step_text(341)) ' vla:2048
    End If
    range("i3") = "ok" ' vla:2050 src:628
    vla_step = 342 ' vla:2051
    If vlatraceon() Then ' vla:2052
        Call vlatracestep(342, vla_step_text(342)) ' vla:2052
    End If
    range("g4") = "West" ' vla:2054 src:629
    vla_step = 343 ' vla:2055
    If vlatraceon() Then ' vla:2056
        Call vlatracestep(343, vla_step_text(343)) ' vla:2056
    End If
    range("h4") = 100 ' vla:2058 src:630
    vla_step = 344 ' vla:2059
    If vlatraceon() Then ' vla:2060
        Call vlatracestep(344, vla_step_text(344)) ' vla:2060
    End If
    range("i4") = "dup" ' vla:2062 src:631
    vla_step = 345 ' vla:2063
    If vlatraceon() Then ' vla:2064
        Call vlatracestep(345, vla_step_text(345)) ' vla:2064
    End If
    Call range("g1:i4").sort(key1:=range("h1"), order1:=xldescending, header:=xlyes)
    vla_step = 346 ' vla:2067
    If vlatraceon() Then ' vla:2068
        Call vlatracestep(346, vla_step_text(346)) ' vla:2068
    End If
    Call range("g1:i4").autofilter(field:=1, criteria1:="West")
    vla_step = 347 ' vla:2071
    If vlatraceon() Then ' vla:2072
        Call vlatracestep(347, vla_step_text(347)) ' vla:2072
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 348 ' vla:2075
    If vlatraceon() Then ' vla:2076
        Call vlatracestep(348, vla_step_text(348)) ' vla:2076
    End If
    Call vlareplaceinrange(range("g1:i4"), "dup", "ok", "values")
    vla_step = 349 ' vla:2079
    If vlatraceon() Then ' vla:2080
        Call vlatracestep(349, vla_step_text(349)) ' vla:2080
    End If
    Call range("g1:i4").removeduplicates(columns:=1, header:=xlyes)
    vla_step = 350 ' vla:2083
    If vlatraceon() Then ' vla:2084
        Call vlatracestep(350, vla_step_text(350)) ' vla:2084
    End If
    Call vlareplaceinrange(columns("i"), "ok", "fine", "values")
    vla_step = 351 ' vla:2087
    If vlatraceon() Then ' vla:2088
        Call vlatracestep(351, vla_step_text(351)) ' vla:2088
    End If
    Call range("g1:i4").copy
    Call range("g1:i4").pastespecial(paste:=xlpastevalues)
    application.cutcopymode = False
    vla_step = 352 ' vla:2091
    If vlatraceon() Then ' vla:2092
        Call vlatracestep(352, vla_step_text(352)) ' vla:2092
    End If
    Call range("g1:i4").clearformats
    vla_step = 353 ' vla:2095
    If vlatraceon() Then ' vla:2096
        Call vlatracestep(353, vla_step_text(353)) ' vla:2096
    End If
    Call vlacheckrangename("demo_table")
    range("g1:i4").name = "demo_table"
    vla_step = 354 ' vla:2099
    If vlatraceon() Then ' vla:2100
        Call vlatracestep(354, vla_step_text(354)) ' vla:2100
    End If
    Call range("g1:i4").copy(destination:=range("k1:m4"))
    vla_step = 355 ' vla:2103
    If vlatraceon() Then ' vla:2104
        Call vlatracestep(355, vla_step_text(355)) ' vla:2104
    End If
    ' ---- instructions.txt:643 Set tab-color of sheet Demo to hot-pink. ----
    worksheets("demo").tab.color = vlacolor(hot_pink)
    vla_step = 356 ' vla:2108
    If vlatraceon() Then ' vla:2109
        Call vlatracestep(356, vla_step_text(356)) ' vla:2109
    End If
    Call activesheet.protect(password:="demo123")
    vla_step = 357 ' vla:2112
    If vlatraceon() Then ' vla:2113
        Call vlatracestep(357, vla_step_text(357)) ' vla:2113
    End If
    Call activesheet.unprotect(password:="demo123")
    vla_step = 358 ' vla:2116
    If vlatraceon() Then ' vla:2117
        Call vlatracestep(358, vla_step_text(358)) ' vla:2117
    End If
    ' ---- instructions.txt:649 Try: ----
    On Error GoTo vla_tryf_4 ' vla:2120 src:675
    vla_step = 359 ' vla:2121 src:675
    If vlatraceon() Then ' vla:2122 src:675
        Call vlatracestep(359, vla_step_text(359)) ' vla:2122 src:675
    End If
    application.displayalerts = False
    Call worksheets("gstruct").delete
    application.displayalerts = True
    GoTo vla_tryd_4 ' vla:2125 src:675
vla_tryf_4: ' vla:2126 src:675
    vla_problem = err.description ' vla:2127 src:675
    Resume vla_tryr_4 ' vla:2128 src:675
vla_tryr_4: ' vla:2129 src:675
    On Error GoTo vla_fail ' vla:2130 src:675
vla_tryd_4: ' vla:2131 src:675
    On Error GoTo vla_fail ' vla:2132 src:675
    vla_step = 360 ' vla:2133
    If vlatraceon() Then ' vla:2134
        Call vlatracestep(360, vla_step_text(360)) ' vla:2134
    End If
    ' ---- instructions.txt:678 Work on sheet GStruct. ----
    Call vlaensuresheet("gstruct") ' vla:2137 src:678
    Call worksheets("gstruct").activate
    vla_step = 361 ' vla:2138
    If vlatraceon() Then ' vla:2139
        Call vlatracestep(361, vla_step_text(361)) ' vla:2139
    End If
    ' ---- instructions.txt:680 Hide row 3. ----
    rows(3).hidden = True
    vla_step = 362 ' vla:2143
    If vlatraceon() Then ' vla:2144
        Call vlatracestep(362, vla_step_text(362)) ' vla:2144
    End If
    columns("b").hidden = True
    vla_step = 363 ' vla:2147
    If vlatraceon() Then ' vla:2148
        Call vlatracestep(363, vla_step_text(363)) ' vla:2148
    End If
    cells.entirerow.hidden = False
    cells.entirecolumn.hidden = False
    vla_step = 364 ' vla:2151
    If vlatraceon() Then ' vla:2152
        Call vlatracestep(364, vla_step_text(364)) ' vla:2152
    End If
    ' ---- instructions.txt:684 Set font size of cell A6 to 36. ----
    range("a6").font.size = 36
    vla_step = 365 ' vla:2156
    If vlatraceon() Then ' vla:2157
        Call vlatracestep(365, vla_step_text(365)) ' vla:2157
    End If
    Call rows(6).autofit
    vla_step = 366 ' vla:2160
    If vlatraceon() Then ' vla:2161
        Call vlatracestep(366, vla_step_text(366)) ' vla:2161
    End If
    ' ---- instructions.txt:687 Group rows 10 through 12. ----
    Call rows((10 & ":" & 12)).group
    vla_step = 367 ' vla:2165
    If vlatraceon() Then ' vla:2166
        Call vlatracestep(367, vla_step_text(367)) ' vla:2166
    End If
    Call rows((14 & ":" & 16)).group
    vla_step = 368 ' vla:2169
    If vlatraceon() Then ' vla:2170
        Call vlatracestep(368, vla_step_text(368)) ' vla:2170
    End If
    Call rows((14 & ":" & 16)).ungroup
    vla_step = 369 ' vla:2173
    If vlatraceon() Then ' vla:2174
        Call vlatracestep(369, vla_step_text(369)) ' vla:2174
    End If
    ' ---- instructions.txt:691 Put "before-insert" into cell A20. ----
    range("a20") = "before-insert" ' vla:2177 src:691
    vla_step = 370 ' vla:2178
    If vlatraceon() Then ' vla:2179
        Call vlatracestep(370, vla_step_text(370)) ' vla:2179
    End If
    Call rows(20).resize(rowsize:=3).insert
    vla_step = 371 ' vla:2182
    If vlatraceon() Then ' vla:2183
        Call vlatracestep(371, vla_step_text(371)) ' vla:2183
    End If
    ' ---- instructions.txt:694 Put "before-delete" into cell A40. ----
    range("a40") = "before-delete" ' vla:2186 src:694
    vla_step = 372 ' vla:2187
    If vlatraceon() Then ' vla:2188
        Call vlatracestep(372, vla_step_text(372)) ' vla:2188
    End If
    Call rows((38 & ":" & 39)).delete
    vla_step = 373 ' vla:2191
    If vlatraceon() Then ' vla:2192
        Call vlatracestep(373, vla_step_text(373)) ' vla:2192
    End If
    ' ---- instructions.txt:697 Put "marker-b" into cell B50. ----
    range("b50") = "marker-b" ' vla:2195 src:697
    vla_step = 374 ' vla:2196
    If vlatraceon() Then ' vla:2197
        Call vlatracestep(374, vla_step_text(374)) ' vla:2197
    End If
    range("c50") = "marker-c" ' vla:2199 src:698
    vla_step = 375 ' vla:2200
    If vlatraceon() Then ' vla:2201
        Call vlatracestep(375, vla_step_text(375)) ' vla:2201
    End If
    range("d50") = "marker-d" ' vla:2203 src:699
    vla_step = 376 ' vla:2204
    If vlatraceon() Then ' vla:2205
        Call vlatracestep(376, vla_step_text(376)) ' vla:2205
    End If
    Call vlamovecolumn("b", "d")
    vla_step = 377 ' vla:2208
    If vlatraceon() Then ' vla:2209
        Call vlatracestep(377, vla_step_text(377)) ' vla:2209
    End If
    ' ---- instructions.txt:702 Freeze the first 2 rows. ----
    Call vlafreezepanes(2)
    vla_step = 378 ' vla:2213
    If vlatraceon() Then ' vla:2214
        Call vlatracestep(378, vla_step_text(378)) ' vla:2214
    End If
    ' ---- instructions.txt:704 Make cell A60 bold. ----
    range("a60").font.bold = True
    vla_step = 379 ' vla:2218
    If vlatraceon() Then ' vla:2219
        Call vlatracestep(379, vla_step_text(379)) ' vla:2219
    End If
    range("a60").interior.color = vbred
    vla_step = 380 ' vla:2222
    If vlatraceon() Then ' vla:2223
        Call vlatracestep(380, vla_step_text(380)) ' vla:2223
    End If
    Call range("a60:a60").copy
    Call range("c60:c60").pastespecial(paste:=xlpasteformats)
    application.cutcopymode = False
    vla_step = 381 ' vla:2226
    If vlatraceon() Then ' vla:2227
        Call vlatracestep(381, vla_step_text(381)) ' vla:2227
    End If
    Call range("a60:a60").copy
    Call range("c62:c62").pastespecial(paste:=xlpasteformats)
    application.cutcopymode = False
    vla_step = 382 ' vla:2230
    If vlatraceon() Then ' vla:2231
        Call vlatracestep(382, vla_step_text(382)) ' vla:2231
    End If
    ' ---- instructions.txt:710 Put formula "=5*2" into cell B64. ----
    range("b64").Formula2 = "=5*2"
    vla_step = 383 ' vla:2235
    If vlatraceon() Then ' vla:2236
        Call vlatracestep(383, vla_step_text(383)) ' vla:2236
    End If
    Call range("b64").copy
    Call range("d64").pastespecial(paste:=xlpasteformulas)
    application.cutcopymode = False
    vla_step = 384 ' vla:2239
    If vlatraceon() Then ' vla:2240
        Call vlatracestep(384, vla_step_text(384)) ' vla:2240
    End If
    ' ---- instructions.txt:713 Set width of column F to 33. ----
    columns("f").columnwidth = 33
    vla_step = 385 ' vla:2244
    If vlatraceon() Then ' vla:2245
        Call vlatracestep(385, vla_step_text(385)) ' vla:2245
    End If
    Call range("f1:f1").copy
    Call range("h1:h1").pastespecial(paste:=xlpastecolumnwidths)
    application.cutcopymode = False
    vla_step = 386 ' vla:2248
    If vlatraceon() Then ' vla:2249
        Call vlatracestep(386, vla_step_text(386)) ' vla:2249
    End If
    ' ---- instructions.txt:716 Put 1 into cell A68. ----
    range("a68") = 1 ' vla:2252 src:716
    vla_step = 387 ' vla:2253
    If vlatraceon() Then ' vla:2254
        Call vlatracestep(387, vla_step_text(387)) ' vla:2254
    End If
    range("a69") = 2 ' vla:2256 src:717
    vla_step = 388 ' vla:2257
    If vlatraceon() Then ' vla:2258
        Call vlatracestep(388, vla_step_text(388)) ' vla:2258
    End If
    range("a70") = 3 ' vla:2260 src:718
    vla_step = 389 ' vla:2261
    If vlatraceon() Then ' vla:2262
        Call vlatracestep(389, vla_step_text(389)) ' vla:2262
    End If
    Call range("a68:a70").copy
    Call range("c68").pastespecial(transpose:=True)
    application.cutcopymode = False
    vla_step = 390 ' vla:2265
    If vlatraceon() Then ' vla:2266
        Call vlatracestep(390, vla_step_text(390)) ' vla:2266
    End If
    ' ---- instructions.txt:721 Put "cutme" into cell A72. ----
    range("a72") = "cutme" ' vla:2269 src:721
    vla_step = 391 ' vla:2270
    If vlatraceon() Then ' vla:2271
        Call vlatracestep(391, vla_step_text(391)) ' vla:2271
    End If
    Call range("a72:a72").cut(destination:=range("c72"))
    vla_step = 392 ' vla:2274
    If vlatraceon() Then ' vla:2275
        Call vlatracestep(392, vla_step_text(392)) ' vla:2275
    End If
    ' ---- instructions.txt:724 Put "rowdata" into cell A74. ----
    range("a74") = "rowdata" ' vla:2278 src:724
    vla_step = 393 ' vla:2279
    If vlatraceon() Then ' vla:2280
        Call vlatracestep(393, vla_step_text(393)) ' vla:2280
    End If
    Call rows(74).copy(destination:=rows(76))
    vla_step = 394 ' vla:2283
    If vlatraceon() Then ' vla:2284
        Call vlatracestep(394, vla_step_text(394)) ' vla:2284
    End If
    ' ---- instructions.txt:727 Put "clearme" into cell A80. ----
    range("a80") = "clearme" ' vla:2287 src:728
    vla_step = 395 ' vla:2288
    If vlatraceon() Then ' vla:2289
        Call vlatracestep(395, vla_step_text(395)) ' vla:2289
    End If
    range("a80").font.bold = True
    vla_step = 396 ' vla:2292
    If vlatraceon() Then ' vla:2293
        Call vlatracestep(396, vla_step_text(396)) ' vla:2293
    End If
    range("a80").interior.color = vbred
    vla_step = 397 ' vla:2296
    If vlatraceon() Then ' vla:2297
        Call vlatracestep(397, vla_step_text(397)) ' vla:2297
    End If
    Call range("a80:b80").clear
    vla_step = 398 ' vla:2300
    If vlatraceon() Then ' vla:2301
        Call vlatracestep(398, vla_step_text(398)) ' vla:2301
    End If
    ' ---- instructions.txt:733 Put "m1" into cell A84. ----
    range("a84") = "m1" ' vla:2304 src:733
    vla_step = 399 ' vla:2305
    If vlatraceon() Then ' vla:2306
        Call vlatracestep(399, vla_step_text(399)) ' vla:2306
    End If
    range("a85") = "m2" ' vla:2308 src:734
    vla_step = 400 ' vla:2309
    If vlatraceon() Then ' vla:2310
        Call vlatracestep(400, vla_step_text(400)) ' vla:2310
    End If
    range("a86") = "m3" ' vla:2312 src:735
    vla_step = 401 ' vla:2313
    If vlatraceon() Then ' vla:2314
        Call vlatracestep(401, vla_step_text(401)) ' vla:2314
    End If
    range("a87") = "m4" ' vla:2316 src:736
    vla_step = 402 ' vla:2317
    If vlatraceon() Then ' vla:2318
        Call vlatracestep(402, vla_step_text(402)) ' vla:2318
    End If
    range("a88") = "m5" ' vla:2320 src:737
    vla_step = 403 ' vla:2321
    If vlatraceon() Then ' vla:2322
        Call vlatracestep(403, vla_step_text(403)) ' vla:2322
    End If
    Call range("a84:a86").delete(shift:=xlshiftup)
    vla_step = 404 ' vla:2325
    If vlatraceon() Then ' vla:2326
        Call vlatracestep(404, vla_step_text(404)) ' vla:2326
    End If
    ' ---- instructions.txt:740 Put "x1" into cell B90. ----
    range("b90") = "x1" ' vla:2329 src:740
    vla_step = 405 ' vla:2330
    If vlatraceon() Then ' vla:2331
        Call vlatracestep(405, vla_step_text(405)) ' vla:2331
    End If
    range("c90") = "x2" ' vla:2333 src:741
    vla_step = 406 ' vla:2334
    If vlatraceon() Then ' vla:2335
        Call vlatracestep(406, vla_step_text(406)) ' vla:2335
    End If
    range("d90") = "x3" ' vla:2337 src:742
    vla_step = 407 ' vla:2338
    If vlatraceon() Then ' vla:2339
        Call vlatracestep(407, vla_step_text(407)) ' vla:2339
    End If
    range("e90") = "rightdata" ' vla:2341 src:743
    vla_step = 408 ' vla:2342
    If vlatraceon() Then ' vla:2343
        Call vlatracestep(408, vla_step_text(408)) ' vla:2343
    End If
    Call range("b90:d90").delete(shift:=xlshifttoleft)
    vla_step = 409 ' vla:2346
    If vlatraceon() Then ' vla:2347
        Call vlatracestep(409, vla_step_text(409)) ' vla:2347
    End If
    ' ---- instructions.txt:746 Put 1 into cell A94. ----
    range("a94") = 1 ' vla:2350 src:746
    vla_step = 410 ' vla:2351
    If vlatraceon() Then ' vla:2352
        Call vlatracestep(410, vla_step_text(410)) ' vla:2352
    End If
    range("b94") = 2 ' vla:2354 src:747
    vla_step = 411 ' vla:2355
    If vlatraceon() Then ' vla:2356
        Call vlatracestep(411, vla_step_text(411)) ' vla:2356
    End If
    range("a95") = 3 ' vla:2358 src:748
    vla_step = 412 ' vla:2359
    If vlatraceon() Then ' vla:2360
        Call vlatracestep(412, vla_step_text(412)) ' vla:2360
    End If
    range("a96") = 5 ' vla:2362 src:749
    vla_step = 413 ' vla:2363
    If vlatraceon() Then ' vla:2364
        Call vlatracestep(413, vla_step_text(413)) ' vla:2364
    End If
    range("b96") = 6 ' vla:2366 src:750
    vla_step = 414 ' vla:2367
    If vlatraceon() Then ' vla:2368
        Call vlatracestep(414, vla_step_text(414)) ' vla:2368
    End If
    Call vladeleteblankrows(range("a94:b96"))
    vla_step = 415 ' vla:2371
    If vlatraceon() Then ' vla:2372
        Call vlatracestep(415, vla_step_text(415)) ' vla:2372
    End If
    ' ---- instructions.txt:753 Put 1 into cell A150. ----
    range("a150") = 1 ' vla:2375 src:754
    vla_step = 416 ' vla:2376
    If vlatraceon() Then ' vla:2377
        Call vlatracestep(416, vla_step_text(416)) ' vla:2377
    End If
    range("a151") = 2 ' vla:2379 src:755
    vla_step = 417 ' vla:2380
    If vlatraceon() Then ' vla:2381
        Call vlatracestep(417, vla_step_text(417)) ' vla:2381
    End If
    range("a152") = 3 ' vla:2383 src:756
    vla_step = 418 ' vla:2384
    If vlatraceon() Then ' vla:2385
        Call vlatracestep(418, vla_step_text(418)) ' vla:2385
    End If
    range("a153") = 4 ' vla:2387 src:757
    vla_step = 419 ' vla:2388
    If vlatraceon() Then ' vla:2389
        Call vlatracestep(419, vla_step_text(419)) ' vla:2389
    End If
    Call vlabandrows(range("a150:d153"), vlacolor("#D9D9D9"))
    vla_step = 420 ' vla:2392
    If vlatraceon() Then ' vla:2393
        Call vlatracestep(420, vla_step_text(420)) ' vla:2393
    End If
    rows(165).font.bold = True
    rows(165).interior.color = vlacolor("#D9D9D9")
    vla_step = 421 ' vla:2396
    If vlatraceon() Then ' vla:2397
        Call vlatracestep(421, vla_step_text(421)) ' vla:2397
    End If
    ' ---- instructions.txt:761 Fill A170:A179 with a series starting at 1. ----
    Call vlafillseries(range("a170:a179"), 1, 1, "linear")
    vla_step = 422 ' vla:2401
    If vlatraceon() Then ' vla:2402
        Call vlatracestep(422, vla_step_text(422)) ' vla:2402
    End If
    Call vlafillseries(range("a180:a184"), 1, 2, "linear")
    vla_step = 423 ' vla:2405
    If vlatraceon() Then ' vla:2406
        Call vlatracestep(423, vla_step_text(423)) ' vla:2406
    End If
    Call vlafillseries(range("a190:a194"), 2, 2, "growth")
    vla_step = 424 ' vla:2409
    If vlatraceon() Then ' vla:2410
        Call vlatracestep(424, vla_step_text(424)) ' vla:2410
    End If
    Call vlafillseries(range("a200:a203"), 2, 3, "growth")
    vla_step = 425 ' vla:2413
    If vlatraceon() Then ' vla:2414
        Call vlatracestep(425, vla_step_text(425)) ' vla:2414
    End If
    ' ---- instructions.txt:767 Make cell Z500 bold. ----
    range("z500").font.bold = True
    vla_step = 426 ' vla:2418
    If vlatraceon() Then ' vla:2419
        Call vlatracestep(426, vla_step_text(426)) ' vla:2419
    End If
    Call range("z500").clear
    vla_step = 427 ' vla:2422
    If vlatraceon() Then ' vla:2423
        Call vlatracestep(427, vla_step_text(427)) ' vla:2423
    End If
    Call vlatrimsheet
    vla_step = 428 ' vla:2426
    If vlatraceon() Then ' vla:2427
        Call vlatracestep(428, vla_step_text(428)) ' vla:2427
    End If
    ' ---- instructions.txt:781 Try: ----
    On Error GoTo vla_tryf_5 ' vla:2430 src:792
    vla_step = 429 ' vla:2431 src:792
    If vlatraceon() Then ' vla:2432 src:792
        Call vlatracestep(429, vla_step_text(429)) ' vla:2432 src:792
    End If
    application.displayalerts = False
    Call worksheets("gformat").delete
    application.displayalerts = True
    GoTo vla_tryd_5 ' vla:2435 src:792
vla_tryf_5: ' vla:2436 src:792
    vla_problem = err.description ' vla:2437 src:792
    Resume vla_tryr_5 ' vla:2438 src:792
vla_tryr_5: ' vla:2439 src:792
    On Error GoTo vla_fail ' vla:2440 src:792
vla_tryd_5: ' vla:2441 src:792
    On Error GoTo vla_fail ' vla:2442 src:792
    vla_step = 430 ' vla:2443
    If vlatraceon() Then ' vla:2444
        Call vlatracestep(430, vla_step_text(430)) ' vla:2444
    End If
    ' ---- instructions.txt:795 Work on sheet GFormat. ----
    Call vlaensuresheet("gformat") ' vla:2447 src:795
    Call worksheets("gformat").activate
    vla_step = 431 ' vla:2448
    If vlatraceon() Then ' vla:2449
        Call vlatracestep(431, vla_step_text(431)) ' vla:2449
    End If
    ' ---- instructions.txt:797 Make range A1:C1 bold. ----
    range("a1:c1").font.bold = True
    vla_step = 432 ' vla:2453
    If vlatraceon() Then ' vla:2454
        Call vlatracestep(432, vla_step_text(432)) ' vla:2454
    End If
    range("a2:c2").font.italic = True
    vla_step = 433 ' vla:2457
    If vlatraceon() Then ' vla:2458
        Call vlatracestep(433, vla_step_text(433)) ' vla:2458
    End If
    range("a3:c3").interior.color = vbblue
    vla_step = 434 ' vla:2461
    If vlatraceon() Then ' vla:2462
        Call vlatracestep(434, vla_step_text(434)) ' vla:2462
    End If
    range("a4:c4").font.color = vlacolor("#FF0000")
    vla_step = 435 ' vla:2465
    If vlatraceon() Then ' vla:2466
        Call vlatracestep(435, vla_step_text(435)) ' vla:2466
    End If
    range("a5:c5").interior.color = vlacolor("#00FF00")
    vla_step = 436 ' vla:2469
    If vlatraceon() Then ' vla:2470
        Call vlatracestep(436, vla_step_text(436)) ' vla:2470
    End If
    range("a6:c6").interior.color = vbyellow
    vla_step = 437 ' vla:2473
    If vlatraceon() Then ' vla:2474
        Call vlatracestep(437, vla_step_text(437)) ' vla:2474
    End If
    range("b6:c6").interior.colorindex = xlnone
    vla_step = 438 ' vla:2477
    If vlatraceon() Then ' vla:2478
        Call vlatracestep(438, vla_step_text(438)) ' vla:2478
    End If
    range("a7:c7").font.size = 16
    vla_step = 439 ' vla:2481
    If vlatraceon() Then ' vla:2482
        Call vlatracestep(439, vla_step_text(439)) ' vla:2482
    End If
    ' ---- instructions.txt:808 Make range A8:C8 bold. ----
    range("a8:c8").font.bold = True
    vla_step = 440 ' vla:2486
    If vlatraceon() Then ' vla:2487
        Call vlatracestep(440, vla_step_text(440)) ' vla:2487
    End If
    range("b8:c8").font.bold = False
    vla_step = 441 ' vla:2490
    If vlatraceon() Then ' vla:2491
        Call vlatracestep(441, vla_step_text(441)) ' vla:2491
    End If
    range("a9:c9").font.italic = True
    vla_step = 442 ' vla:2494
    If vlatraceon() Then ' vla:2495
        Call vlatracestep(442, vla_step_text(442)) ' vla:2495
    End If
    range("c9").font.italic = False
    vla_step = 443 ' vla:2498
    If vlatraceon() Then ' vla:2499
        Call vlatracestep(443, vla_step_text(443)) ' vla:2499
    End If
    range("a10:c10").font.underline = xlunderlinestylesingle
    vla_step = 444 ' vla:2502
    If vlatraceon() Then ' vla:2503
        Call vlatracestep(444, vla_step_text(444)) ' vla:2503
    End If
    range("a11:c11").font.strikethrough = True
    vla_step = 445 ' vla:2506
    If vlatraceon() Then ' vla:2507
        Call vlatracestep(445, vla_step_text(445)) ' vla:2507
    End If
    range("a12:c12").font.underline = xlunderlinestylesingle
    vla_step = 446 ' vla:2510
    If vlatraceon() Then ' vla:2511
        Call vlatracestep(446, vla_step_text(446)) ' vla:2511
    End If
    range("b12:c12").font.underline = xlunderlinestylenone
    vla_step = 447 ' vla:2514
    If vlatraceon() Then ' vla:2515
        Call vlatracestep(447, vla_step_text(447)) ' vla:2515
    End If
    range("a13:c13").font.strikethrough = True
    vla_step = 448 ' vla:2518
    If vlatraceon() Then ' vla:2519
        Call vlatracestep(448, vla_step_text(448)) ' vla:2519
    End If
    range("c13").font.strikethrough = False
    vla_step = 449 ' vla:2522
    If vlatraceon() Then ' vla:2523
        Call vlatracestep(449, vla_step_text(449)) ' vla:2523
    End If
    range("a14:c14").font.name = "Courier New"
    vla_step = 450 ' vla:2526
    If vlatraceon() Then ' vla:2527
        Call vlatracestep(450, vla_step_text(450)) ' vla:2527
    End If
    ' ---- instructions.txt:821 Align range A15:C15 to the top. ----
    range("a15:c15").verticalalignment = xltop
    vla_step = 451 ' vla:2531
    If vlatraceon() Then ' vla:2532
        Call vlatracestep(451, vla_step_text(451)) ' vla:2532
    End If
    range("a16:c16").verticalalignment = xlcenter
    vla_step = 452 ' vla:2535
    If vlatraceon() Then ' vla:2536
        Call vlatracestep(452, vla_step_text(452)) ' vla:2536
    End If
    range("a17:c17").verticalalignment = xltop
    vla_step = 453 ' vla:2539
    If vlatraceon() Then ' vla:2540
        Call vlatracestep(453, vla_step_text(453)) ' vla:2540
    End If
    range("b17:c17").verticalalignment = xlbottom
    vla_step = 454 ' vla:2543
    If vlatraceon() Then ' vla:2544
        Call vlatracestep(454, vla_step_text(454)) ' vla:2544
    End If
    range("a18:c18").indentlevel = 2
    vla_step = 455 ' vla:2547
    If vlatraceon() Then ' vla:2548
        Call vlatracestep(455, vla_step_text(455)) ' vla:2548
    End If
    range("a19:c19").orientation = 45
    vla_step = 456 ' vla:2551
    If vlatraceon() Then ' vla:2552
        Call vlatracestep(456, vla_step_text(456)) ' vla:2552
    End If
    ' ---- instructions.txt:830 Add a border around range B21:D23. ----
    range("b21:d23").borders(xledgetop).linestyle = xlcontinuous
    range("b21:d23").borders(xledgebottom).linestyle = xlcontinuous
    range("b21:d23").borders(xledgeleft).linestyle = xlcontinuous
    range("b21:d23").borders(xledgeright).linestyle = xlcontinuous
    vla_step = 457 ' vla:2556
    If vlatraceon() Then ' vla:2557
        Call vlatracestep(457, vla_step_text(457)) ' vla:2557
    End If
    range("b25:d25").borders(xledgebottom).linestyle = xlcontinuous
    vla_step = 458 ' vla:2560
    If vlatraceon() Then ' vla:2561
        Call vlatracestep(458, vla_step_text(458)) ' vla:2561
    End If
    range("b27:d29").borders.linestyle = xlcontinuous
    vla_step = 459 ' vla:2564
    If vlatraceon() Then ' vla:2565
        Call vlatracestep(459, vla_step_text(459)) ' vla:2565
    End If
    range("b31:d33").borders.linestyle = xlcontinuous
    vla_step = 460 ' vla:2568
    If vlatraceon() Then ' vla:2569
        Call vlatracestep(460, vla_step_text(460)) ' vla:2569
    End If
    range("c31:d33").borders.linestyle = xlnone
    vla_step = 461 ' vla:2572
    If vlatraceon() Then ' vla:2573
        Call vlatracestep(461, vla_step_text(461)) ' vla:2573
    End If
    ' ---- instructions.txt:840 Set width of column F to 40. ----
    columns("f").columnwidth = 40
    vla_step = 462 ' vla:2577
    If vlatraceon() Then ' vla:2578
        Call vlatracestep(462, vla_step_text(462)) ' vla:2578
    End If
    range("f40") = 1234.56 ' vla:2580 src:847
    vla_step = 463 ' vla:2581
    If vlatraceon() Then ' vla:2582
        Call vlatracestep(463, vla_step_text(463)) ' vla:2582
    End If
    range("f40").numberformat = vlanumberformatcode("number", 2)
    vla_step = 464 ' vla:2585
    If vlatraceon() Then ' vla:2586
        Call vlatracestep(464, vla_step_text(464)) ' vla:2586
    End If
    range("f41") = 1234.56 ' vla:2588 src:849
    vla_step = 465 ' vla:2589
    If vlatraceon() Then ' vla:2590
        Call vlatracestep(465, vla_step_text(465)) ' vla:2590
    End If
    range("f41").numberformat = vlanumberformatcode("number", 3)
    vla_step = 466 ' vla:2593
    If vlatraceon() Then ' vla:2594
        Call vlatracestep(466, vla_step_text(466)) ' vla:2594
    End If
    range("f42") = 1234.56 ' vla:2596 src:851
    vla_step = 467 ' vla:2597
    If vlatraceon() Then ' vla:2598
        Call vlatracestep(467, vla_step_text(467)) ' vla:2598
    End If
    range("f42").numberformat = vlanumberformatcode("number-separated", 2)
    vla_step = 468 ' vla:2601
    If vlatraceon() Then ' vla:2602
        Call vlatracestep(468, vla_step_text(468)) ' vla:2602
    End If
    range("f43") = 1234.56 ' vla:2604 src:853
    vla_step = 469 ' vla:2605
    If vlatraceon() Then ' vla:2606
        Call vlatracestep(469, vla_step_text(469)) ' vla:2606
    End If
    range("f43").numberformat = vlanumberformatcode("number-separated", 0)
    vla_step = 470 ' vla:2609
    If vlatraceon() Then ' vla:2610
        Call vlatracestep(470, vla_step_text(470)) ' vla:2610
    End If
    range("f44") = 1234.56 ' vla:2612 src:855
    vla_step = 471 ' vla:2613
    If vlatraceon() Then ' vla:2614
        Call vlatracestep(471, vla_step_text(471)) ' vla:2614
    End If
    range("f44").numberformat = vlanumberformatcode("dollars", 2)
    vla_step = 472 ' vla:2617
    If vlatraceon() Then ' vla:2618
        Call vlatracestep(472, vla_step_text(472)) ' vla:2618
    End If
    range("f45") = 1234.56 ' vla:2620 src:857
    vla_step = 473 ' vla:2621
    If vlatraceon() Then ' vla:2622
        Call vlatracestep(473, vla_step_text(473)) ' vla:2622
    End If
    range("f45").numberformat = vlanumberformatcode("euros", 0)
    vla_step = 474 ' vla:2625
    If vlatraceon() Then ' vla:2626
        Call vlatracestep(474, vla_step_text(474)) ' vla:2626
    End If
    range("f46") = 1234.56 ' vla:2628 src:859
    vla_step = 475 ' vla:2629
    If vlatraceon() Then ' vla:2630
        Call vlatracestep(475, vla_step_text(475)) ' vla:2630
    End If
    range("f46").numberformat = vlanumberformatcode("pounds", 2)
    vla_step = 476 ' vla:2633
    If vlatraceon() Then ' vla:2634
        Call vlatracestep(476, vla_step_text(476)) ' vla:2634
    End If
    range("f47") = 1234.56 ' vla:2636 src:861
    vla_step = 477 ' vla:2637
    If vlatraceon() Then ' vla:2638
        Call vlatracestep(477, vla_step_text(477)) ' vla:2638
    End If
    range("f47").numberformat = vlanumberformatcode("accounting-dollars", 2)
    vla_step = 478 ' vla:2641
    If vlatraceon() Then ' vla:2642
        Call vlatracestep(478, vla_step_text(478)) ' vla:2642
    End If
    range("f48") = -1234.56 ' vla:2644 src:863
    vla_step = 479 ' vla:2645
    If vlatraceon() Then ' vla:2646
        Call vlatracestep(479, vla_step_text(479)) ' vla:2646
    End If
    range("f48").numberformat = vlanumberformatcode("accounting-euros", 0)
    vla_step = 480 ' vla:2649
    If vlatraceon() Then ' vla:2650
        Call vlatracestep(480, vla_step_text(480)) ' vla:2650
    End If
    range("f49") = 0.125 ' vla:2652 src:865
    vla_step = 481 ' vla:2653
    If vlatraceon() Then ' vla:2654
        Call vlatracestep(481, vla_step_text(481)) ' vla:2654
    End If
    range("f49").numberformat = "0.0%"
    vla_step = 482 ' vla:2657
    If vlatraceon() Then ' vla:2658
        Call vlatracestep(482, vla_step_text(482)) ' vla:2658
    End If
    range("f50") = 0.125 ' vla:2660 src:867
    vla_step = 483 ' vla:2661
    If vlatraceon() Then ' vla:2662
        Call vlatracestep(483, vla_step_text(483)) ' vla:2662
    End If
    range("f50").numberformat = vlanumberformatcode("percent", 2)
    vla_step = 484 ' vla:2665
    If vlatraceon() Then ' vla:2666
        Call vlatracestep(484, vla_step_text(484)) ' vla:2666
    End If
    range("f51") = 46000 ' vla:2668 src:869
    vla_step = 485 ' vla:2669
    If vlatraceon() Then ' vla:2670
        Call vlatracestep(485, vla_step_text(485)) ' vla:2670
    End If
    range("f51").numberformat = "m/d/yyyy"
    vla_step = 486 ' vla:2673
    If vlatraceon() Then ' vla:2674
        Call vlatracestep(486, vla_step_text(486)) ' vla:2674
    End If
    range("f52") = 46000 ' vla:2676 src:871
    vla_step = 487 ' vla:2677
    If vlatraceon() Then ' vla:2678
        Call vlatracestep(487, vla_step_text(487)) ' vla:2678
    End If
    range("f52").numberformat = "[$-F800]dddd, mmmm dd, yyyy"
    vla_step = 488 ' vla:2681
    If vlatraceon() Then ' vla:2682
        Call vlatracestep(488, vla_step_text(488)) ' vla:2682
    End If
    range("f53") = 46000 ' vla:2684 src:873
    vla_step = 489 ' vla:2685
    If vlatraceon() Then ' vla:2686
        Call vlatracestep(489, vla_step_text(489)) ' vla:2686
    End If
    range("f53").numberformat = "yyyy-mm-dd"
    vla_step = 490 ' vla:2689
    If vlatraceon() Then ' vla:2690
        Call vlatracestep(490, vla_step_text(490)) ' vla:2690
    End If
    range("f54") = 0.5625 ' vla:2692 src:875
    vla_step = 491 ' vla:2693
    If vlatraceon() Then ' vla:2694
        Call vlatracestep(491, vla_step_text(491)) ' vla:2694
    End If
    range("f54").numberformat = "[$-F400]h:mm:ss AM/PM"
    vla_step = 492 ' vla:2697
    If vlatraceon() Then ' vla:2698
        Call vlatracestep(492, vla_step_text(492)) ' vla:2698
    End If
    range("f55") = 42 ' vla:2700 src:877
    vla_step = 493 ' vla:2701
    If vlatraceon() Then ' vla:2702
        Call vlatracestep(493, vla_step_text(493)) ' vla:2702
    End If
    range("f55").numberformat = "@"
    vla_step = 494 ' vla:2705
    If vlatraceon() Then ' vla:2706
        Call vlatracestep(494, vla_step_text(494)) ' vla:2706
    End If
    range("f56") = 1234.56 ' vla:2708 src:879
    vla_step = 495 ' vla:2709
    If vlatraceon() Then ' vla:2710
        Call vlatracestep(495, vla_step_text(495)) ' vla:2710
    End If
    range("f56").numberformat = vlanumberformatcode("dollars", 2)
    vla_step = 496 ' vla:2713
    If vlatraceon() Then ' vla:2714
        Call vlatracestep(496, vla_step_text(496)) ' vla:2714
    End If
    range("f56").numberformat = "General"
    vla_step = 497 ' vla:2717
    If vlatraceon() Then ' vla:2718
        Call vlatracestep(497, vla_step_text(497)) ' vla:2718
    End If
    range("f57") = 42 ' vla:2720 src:882
    vla_step = 498 ' vla:2721
    If vlatraceon() Then ' vla:2722
        Call vlatracestep(498, vla_step_text(498)) ' vla:2722
    End If
    range("f57").numberformat = "00000"
    vla_step = 499 ' vla:2725
    If vlatraceon() Then ' vla:2726
        Call vlatracestep(499, vla_step_text(499)) ' vla:2726
    End If
    ' ---- instructions.txt:885 Add a border colored "#FF0000" around range H40:J42. ----
    range("h40:j42").borders(xledgetop).linestyle = xlcontinuous
    range("h40:j42").borders(xledgetop).color = vlacolor("#FF0000")
    range("h40:j42").borders(xledgebottom).linestyle = xlcontinuous
    range("h40:j42").borders(xledgebottom).color = vlacolor("#FF0000")
    range("h40:j42").borders(xledgeleft).linestyle = xlcontinuous
    range("h40:j42").borders(xledgeleft).color = vlacolor("#FF0000")
    range("h40:j42").borders(xledgeright).linestyle = xlcontinuous
    range("h40:j42").borders(xledgeright).color = vlacolor("#FF0000")
    vla_step = 500 ' vla:2730
    If vlatraceon() Then ' vla:2731
        Call vlatracestep(500, vla_step_text(500)) ' vla:2731
    End If
    range("h44:j44").borders(xledgebottom).linestyle = xlcontinuous
    range("h44:j44").borders(xledgebottom).color = vlacolor("green")
    vla_step = 501 ' vla:2734
    If vlatraceon() Then ' vla:2735
        Call vlatracestep(501, vla_step_text(501)) ' vla:2735
    End If
    range("h46:j48").borders.linestyle = xlcontinuous
    range("h46:j48").borders.color = vlacolor("blue")
    vla_step = 502 ' vla:2738
    If vlatraceon() Then ' vla:2739
        Call vlatracestep(502, vla_step_text(502)) ' vla:2739
    End If
    ' ---- instructions.txt:893 Try: ----
    On Error GoTo vla_tryf_6 ' vla:2742 src:899
    vla_step = 503 ' vla:2743 src:899
    If vlatraceon() Then ' vla:2744 src:899
        Call vlatracestep(503, vla_step_text(503)) ' vla:2744 src:899
    End If
    application.displayalerts = False
    Call worksheets("gsortfilter").delete
    application.displayalerts = True
    GoTo vla_tryd_6 ' vla:2747 src:899
vla_tryf_6: ' vla:2748 src:899
    vla_problem = err.description ' vla:2749 src:899
    Resume vla_tryr_6 ' vla:2750 src:899
vla_tryr_6: ' vla:2751 src:899
    On Error GoTo vla_fail ' vla:2752 src:899
vla_tryd_6: ' vla:2753 src:899
    On Error GoTo vla_fail ' vla:2754 src:899
    vla_step = 504 ' vla:2755
    If vlatraceon() Then ' vla:2756
        Call vlatracestep(504, vla_step_text(504)) ' vla:2756
    End If
    ' ---- instructions.txt:902 Work on sheet GSortFilter. ----
    Call vlaensuresheet("gsortfilter") ' vla:2759 src:902
    Call worksheets("gsortfilter").activate
    vla_step = 505 ' vla:2760
    If vlatraceon() Then ' vla:2761
        Call vlatracestep(505, vla_step_text(505)) ' vla:2761
    End If
    ' ---- instructions.txt:904 Put "Item" into cell A1. ----
    range("a1") = "Item" ' vla:2764 src:906
    vla_step = 506 ' vla:2765
    If vlatraceon() Then ' vla:2766
        Call vlatracestep(506, vla_step_text(506)) ' vla:2766
    End If
    range("b1") = "Qty" ' vla:2768 src:907
    vla_step = 507 ' vla:2769
    If vlatraceon() Then ' vla:2770
        Call vlatracestep(507, vla_step_text(507)) ' vla:2770
    End If
    range("a2") = "a" ' vla:2772 src:908
    vla_step = 508 ' vla:2773
    If vlatraceon() Then ' vla:2774
        Call vlatracestep(508, vla_step_text(508)) ' vla:2774
    End If
    range("b2") = 2 ' vla:2776 src:909
    vla_step = 509 ' vla:2777
    If vlatraceon() Then ' vla:2778
        Call vlatracestep(509, vla_step_text(509)) ' vla:2778
    End If
    range("a3") = "b" ' vla:2780 src:910
    vla_step = 510 ' vla:2781
    If vlatraceon() Then ' vla:2782
        Call vlatracestep(510, vla_step_text(510)) ' vla:2782
    End If
    range("b3") = 9 ' vla:2784 src:911
    vla_step = 511 ' vla:2785
    If vlatraceon() Then ' vla:2786
        Call vlatracestep(511, vla_step_text(511)) ' vla:2786
    End If
    range("a4") = "c" ' vla:2788 src:912
    vla_step = 512 ' vla:2789
    If vlatraceon() Then ' vla:2790
        Call vlatracestep(512, vla_step_text(512)) ' vla:2790
    End If
    range("b4") = 5 ' vla:2792 src:913
    vla_step = 513 ' vla:2793
    If vlatraceon() Then ' vla:2794
        Call vlatracestep(513, vla_step_text(513)) ' vla:2794
    End If
    Call activesheet.usedrange.sort(key1:=activesheet.usedrange.columns(vlacolumninrange(activesheet.usedrange, "b")), order1:=xldescending, header:=xlyes)
    vla_step = 514 ' vla:2797
    If vlatraceon() Then ' vla:2798
        Call vlatracestep(514, vla_step_text(514)) ' vla:2798
    End If
    ' ---- instructions.txt:916 Put "Name" into cell D1. ----
    range("d1") = "Name" ' vla:2801 src:918
    vla_step = 515 ' vla:2802
    If vlatraceon() Then ' vla:2803
        Call vlatracestep(515, vla_step_text(515)) ' vla:2803
    End If
    range("e1") = "Score" ' vla:2805 src:919
    vla_step = 516 ' vla:2806
    If vlatraceon() Then ' vla:2807
        Call vlatracestep(516, vla_step_text(516)) ' vla:2807
    End If
    range("d2") = "p" ' vla:2809 src:920
    vla_step = 517 ' vla:2810
    If vlatraceon() Then ' vla:2811
        Call vlatracestep(517, vla_step_text(517)) ' vla:2811
    End If
    range("e2") = 3 ' vla:2813 src:921
    vla_step = 518 ' vla:2814
    If vlatraceon() Then ' vla:2815
        Call vlatracestep(518, vla_step_text(518)) ' vla:2815
    End If
    range("d3") = "q" ' vla:2817 src:922
    vla_step = 519 ' vla:2818
    If vlatraceon() Then ' vla:2819
        Call vlatracestep(519, vla_step_text(519)) ' vla:2819
    End If
    range("e3") = 1 ' vla:2821 src:923
    vla_step = 520 ' vla:2822
    If vlatraceon() Then ' vla:2823
        Call vlatracestep(520, vla_step_text(520)) ' vla:2823
    End If
    range("d4") = "r" ' vla:2825 src:924
    vla_step = 521 ' vla:2826
    If vlatraceon() Then ' vla:2827
        Call vlatracestep(521, vla_step_text(521)) ' vla:2827
    End If
    range("e4") = 2 ' vla:2829 src:925
    vla_step = 522 ' vla:2830
    If vlatraceon() Then ' vla:2831
        Call vlatracestep(522, vla_step_text(522)) ' vla:2831
    End If
    Call range("d1:e4").sort(key1:=range("d1:e4").columns(vlacolumninrange(range("d1:e4"), "e")), order1:=xlascending, header:=xlyes)
    vla_step = 523 ' vla:2834
    If vlatraceon() Then ' vla:2835
        Call vlatracestep(523, vla_step_text(523)) ' vla:2835
    End If
    ' ---- instructions.txt:928 Put "x" into cell G2. ----
    range("g2") = "x" ' vla:2838 src:929
    vla_step = 524 ' vla:2839
    If vlatraceon() Then ' vla:2840
        Call vlatracestep(524, vla_step_text(524)) ' vla:2840
    End If
    range("h2") = 5 ' vla:2842 src:930
    vla_step = 525 ' vla:2843
    If vlatraceon() Then ' vla:2844
        Call vlatracestep(525, vla_step_text(525)) ' vla:2844
    End If
    range("g3") = "y" ' vla:2846 src:931
    vla_step = 526 ' vla:2847
    If vlatraceon() Then ' vla:2848
        Call vlatracestep(526, vla_step_text(526)) ' vla:2848
    End If
    range("h3") = 9 ' vla:2850 src:932
    vla_step = 527 ' vla:2851
    If vlatraceon() Then ' vla:2852
        Call vlatracestep(527, vla_step_text(527)) ' vla:2852
    End If
    range("g4") = "z" ' vla:2854 src:933
    vla_step = 528 ' vla:2855
    If vlatraceon() Then ' vla:2856
        Call vlatracestep(528, vla_step_text(528)) ' vla:2856
    End If
    range("h4") = 7 ' vla:2858 src:934
    vla_step = 529 ' vla:2859
    If vlatraceon() Then ' vla:2860
        Call vlatracestep(529, vla_step_text(529)) ' vla:2860
    End If
    Call range("g2:h4").sort(key1:=range("g2:h4").columns(vlacolumninrange(range("g2:h4"), "h")), order1:=xldescending, header:=xlno)
    vla_step = 530 ' vla:2863
    If vlatraceon() Then ' vla:2864
        Call vlatracestep(530, vla_step_text(530)) ' vla:2864
    End If
    ' ---- instructions.txt:937 Put "Region" into cell J1. ----
    range("j1") = "Region" ' vla:2867 src:939
    vla_step = 531 ' vla:2868
    If vlatraceon() Then ' vla:2869
        Call vlatracestep(531, vla_step_text(531)) ' vla:2869
    End If
    range("k1") = "Amount" ' vla:2871 src:940
    vla_step = 532 ' vla:2872
    If vlatraceon() Then ' vla:2873
        Call vlatracestep(532, vla_step_text(532)) ' vla:2873
    End If
    range("l1") = "Tag" ' vla:2875 src:941
    vla_step = 533 ' vla:2876
    If vlatraceon() Then ' vla:2877
        Call vlatracestep(533, vla_step_text(533)) ' vla:2877
    End If
    range("j2") = "West" ' vla:2879 src:942
    vla_step = 534 ' vla:2880
    If vlatraceon() Then ' vla:2881
        Call vlatracestep(534, vla_step_text(534)) ' vla:2881
    End If
    range("k2") = 100 ' vla:2883 src:943
    vla_step = 535 ' vla:2884
    If vlatraceon() Then ' vla:2885
        Call vlatracestep(535, vla_step_text(535)) ' vla:2885
    End If
    range("l2") = "a" ' vla:2887 src:944
    vla_step = 536 ' vla:2888
    If vlatraceon() Then ' vla:2889
        Call vlatracestep(536, vla_step_text(536)) ' vla:2889
    End If
    range("j3") = "East" ' vla:2891 src:945
    vla_step = 537 ' vla:2892
    If vlatraceon() Then ' vla:2893
        Call vlatracestep(537, vla_step_text(537)) ' vla:2893
    End If
    range("k3") = 50 ' vla:2895 src:946
    vla_step = 538 ' vla:2896
    If vlatraceon() Then ' vla:2897
        Call vlatracestep(538, vla_step_text(538)) ' vla:2897
    End If
    range("l3") = "b" ' vla:2899 src:947
    vla_step = 539 ' vla:2900
    If vlatraceon() Then ' vla:2901
        Call vlatracestep(539, vla_step_text(539)) ' vla:2901
    End If
    range("j4") = "West" ' vla:2903 src:948
    vla_step = 540 ' vla:2904
    If vlatraceon() Then ' vla:2905
        Call vlatracestep(540, vla_step_text(540)) ' vla:2905
    End If
    range("k4") = 300 ' vla:2907 src:949
    vla_step = 541 ' vla:2908
    If vlatraceon() Then ' vla:2909
        Call vlatracestep(541, vla_step_text(541)) ' vla:2909
    End If
    range("l4") = "c" ' vla:2911 src:950
    vla_step = 542 ' vla:2912
    If vlatraceon() Then ' vla:2913
        Call vlatracestep(542, vla_step_text(542)) ' vla:2913
    End If
    range("j5") = "East" ' vla:2915 src:951
    vla_step = 543 ' vla:2916
    If vlatraceon() Then ' vla:2917
        Call vlatracestep(543, vla_step_text(543)) ' vla:2917
    End If
    range("k5") = 250 ' vla:2919 src:952
    vla_step = 544 ' vla:2920
    If vlatraceon() Then ' vla:2921
        Call vlatracestep(544, vla_step_text(544)) ' vla:2921
    End If
    range("l5") = "d" ' vla:2923 src:953
    vla_step = 545 ' vla:2924
    If vlatraceon() Then ' vla:2925
        Call vlatracestep(545, vla_step_text(545)) ' vla:2925
    End If
    Call range("j1:l5").sort(key1:=range("j1:l5").columns(vlacolumninrange(range("j1:l5"), "j")), order1:=xlascending, key2:=range("j1:l5").columns(vlacolumninrange(range("j1:l5"), "k")), order2:=xldescending, header:=xlyes)
    vla_step = 546 ' vla:2928
    If vlatraceon() Then ' vla:2929
        Call vlatracestep(546, vla_step_text(546)) ' vla:2929
    End If
    ' ---- instructions.txt:956 Put 2 into cell N2. ----
    range("n2") = 2 ' vla:2932 src:957
    vla_step = 547 ' vla:2933
    If vlatraceon() Then ' vla:2934
        Call vlatracestep(547, vla_step_text(547)) ' vla:2934
    End If
    range("o2") = 9 ' vla:2936 src:958
    vla_step = 548 ' vla:2937
    If vlatraceon() Then ' vla:2938
        Call vlatracestep(548, vla_step_text(548)) ' vla:2938
    End If
    range("n3") = 1 ' vla:2940 src:959
    vla_step = 549 ' vla:2941
    If vlatraceon() Then ' vla:2942
        Call vlatracestep(549, vla_step_text(549)) ' vla:2942
    End If
    range("o3") = 8 ' vla:2944 src:960
    vla_step = 550 ' vla:2945
    If vlatraceon() Then ' vla:2946
        Call vlatracestep(550, vla_step_text(550)) ' vla:2946
    End If
    range("n4") = 2 ' vla:2948 src:961
    vla_step = 551 ' vla:2949
    If vlatraceon() Then ' vla:2950
        Call vlatracestep(551, vla_step_text(551)) ' vla:2950
    End If
    range("o4") = 7 ' vla:2952 src:962
    vla_step = 552 ' vla:2953
    If vlatraceon() Then ' vla:2954
        Call vlatracestep(552, vla_step_text(552)) ' vla:2954
    End If
    Call range("n2:o4").sort(key1:=range("n2:o4").columns(vlacolumninrange(range("n2:o4"), "n")), order1:=xlascending, key2:=range("n2:o4").columns(vlacolumninrange(range("n2:o4"), "o")), order2:=xlascending, header:=xlno)
    vla_step = 553 ' vla:2957
    If vlatraceon() Then ' vla:2958
        Call vlatracestep(553, vla_step_text(553)) ' vla:2958
    End If
    ' ---- instructions.txt:965 Put "Region" into cell A40. ----
    range("a40") = "Region" ' vla:2961 src:972
    vla_step = 554 ' vla:2962
    If vlatraceon() Then ' vla:2963
        Call vlatracestep(554, vla_step_text(554)) ' vla:2963
    End If
    range("b40") = "Amount" ' vla:2965 src:973
    vla_step = 555 ' vla:2966
    If vlatraceon() Then ' vla:2967
        Call vlatracestep(555, vla_step_text(555)) ' vla:2967
    End If
    range("a41") = "West" ' vla:2969 src:974
    vla_step = 556 ' vla:2970
    If vlatraceon() Then ' vla:2971
        Call vlatracestep(556, vla_step_text(556)) ' vla:2971
    End If
    range("b41") = 1 ' vla:2973 src:975
    vla_step = 557 ' vla:2974
    If vlatraceon() Then ' vla:2975
        Call vlatracestep(557, vla_step_text(557)) ' vla:2975
    End If
    range("a42") = "East" ' vla:2977 src:976
    vla_step = 558 ' vla:2978
    If vlatraceon() Then ' vla:2979
        Call vlatracestep(558, vla_step_text(558)) ' vla:2979
    End If
    range("b42") = 2 ' vla:2981 src:977
    vla_step = 559 ' vla:2982
    If vlatraceon() Then ' vla:2983
        Call vlatracestep(559, vla_step_text(559)) ' vla:2983
    End If
    range("a43") = "West" ' vla:2985 src:978
    vla_step = 560 ' vla:2986
    If vlatraceon() Then ' vla:2987
        Call vlatracestep(560, vla_step_text(560)) ' vla:2987
    End If
    range("b43") = 3 ' vla:2989 src:979
    vla_step = 561 ' vla:2990
    If vlatraceon() Then ' vla:2991
        Call vlatracestep(561, vla_step_text(561)) ' vla:2991
    End If
    Call range("a40:b43").autofilter(field:=vlafilterfield(range("a40:b43"), "a"), criteria1:=vlafiltercriterion("at-least", "West"), operator:=xland, criteria2:=vlafiltercriterion("at-most", "West"))
    vla_step = 562 ' vla:2994
    If vlatraceon() Then ' vla:2995
        Call vlatracestep(562, vla_step_text(562)) ' vla:2995
    End If
    Call range("a40:b43").specialcells(xlcelltypevisible).copy(destination:=range("d46"))
    vla_step = 563 ' vla:2998
    If vlatraceon() Then ' vla:2999
        Call vlatracestep(563, vla_step_text(563)) ' vla:2999
    End If
    activesheet.autofiltermode = False
    vla_step = 564 ' vla:3002
    If vlatraceon() Then ' vla:3003
        Call vlatracestep(564, vla_step_text(564)) ' vla:3003
    End If
    ' ---- instructions.txt:984 Try: ----
    On Error GoTo vla_tryf_7 ' vla:3006 src:988
    vla_step = 565 ' vla:3007 src:988
    If vlatraceon() Then ' vla:3008 src:988
        Call vlatracestep(565, vla_step_text(565)) ' vla:3008 src:988
    End If
    Call range("a40:b43").autofilter(field:=vlafilterfield(range("a40:b43"), "b"), criteria1:=vlafiltercriterion("greater", "abc"))
    GoTo vla_tryd_7 ' vla:3011 src:988
vla_tryf_7: ' vla:3012 src:988
    vla_problem = err.description ' vla:3013 src:988
    Resume vla_tryr_7 ' vla:3014 src:988
vla_tryr_7: ' vla:3015 src:988
    On Error GoTo vla_fail ' vla:3016 src:988
vla_tryd_7: ' vla:3017 src:988
    On Error GoTo vla_fail ' vla:3018 src:988
    vla_step = 566 ' vla:3019
    If vlatraceon() Then ' vla:3020
        Call vlatracestep(566, vla_step_text(566)) ' vla:3020
    End If
    ' ---- instructions.txt:991 Put "Region" into cell A20. ----
    range("a20") = "Region" ' vla:3023 src:994
    vla_step = 567 ' vla:3024
    If vlatraceon() Then ' vla:3025
        Call vlatracestep(567, vla_step_text(567)) ' vla:3025
    End If
    range("b20") = "Amount" ' vla:3027 src:995
    vla_step = 568 ' vla:3028
    If vlatraceon() Then ' vla:3029
        Call vlatracestep(568, vla_step_text(568)) ' vla:3029
    End If
    range("c20") = "Code" ' vla:3031 src:996
    vla_step = 569 ' vla:3032
    If vlatraceon() Then ' vla:3033
        Call vlatracestep(569, vla_step_text(569)) ' vla:3033
    End If
    range("a21") = "West" ' vla:3035 src:997
    vla_step = 570 ' vla:3036
    If vlatraceon() Then ' vla:3037
        Call vlatracestep(570, vla_step_text(570)) ' vla:3037
    End If
    range("b21") = 100 ' vla:3039 src:998
    vla_step = 571 ' vla:3040
    If vlatraceon() Then ' vla:3041
        Call vlatracestep(571, vla_step_text(571)) ' vla:3041
    End If
    range("c21") = "5*3" ' vla:3043 src:999
    vla_step = 572 ' vla:3044
    If vlatraceon() Then ' vla:3045
        Call vlatracestep(572, vla_step_text(572)) ' vla:3045
    End If
    range("a22") = "East" ' vla:3047 src:1000
    vla_step = 573 ' vla:3048
    If vlatraceon() Then ' vla:3049
        Call vlatracestep(573, vla_step_text(573)) ' vla:3049
    End If
    range("b22") = 250 ' vla:3051 src:1001
    vla_step = 574 ' vla:3052
    If vlatraceon() Then ' vla:3053
        Call vlatracestep(574, vla_step_text(574)) ' vla:3053
    End If
    range("c22") = "53" ' vla:3055 src:1002
    vla_step = 575 ' vla:3056
    If vlatraceon() Then ' vla:3057
        Call vlatracestep(575, vla_step_text(575)) ' vla:3057
    End If
    range("a23") = "Western" ' vla:3059 src:1003
    vla_step = 576 ' vla:3060
    If vlatraceon() Then ' vla:3061
        Call vlatracestep(576, vla_step_text(576)) ' vla:3061
    End If
    range("b23") = 100 ' vla:3063 src:1004
    vla_step = 577 ' vla:3064
    If vlatraceon() Then ' vla:3065
        Call vlatracestep(577, vla_step_text(577)) ' vla:3065
    End If
    range("c23") = "x" ' vla:3067 src:1005
    vla_step = 578 ' vla:3068
    If vlatraceon() Then ' vla:3069
        Call vlatracestep(578, vla_step_text(578)) ' vla:3069
    End If
    range("a24") = "West" ' vla:3071 src:1006
    vla_step = 579 ' vla:3072
    If vlatraceon() Then ' vla:3073
        Call vlatracestep(579, vla_step_text(579)) ' vla:3073
    End If
    range("b24") = 50 ' vla:3075 src:1007
    vla_step = 580 ' vla:3076
    If vlatraceon() Then ' vla:3077
        Call vlatracestep(580, vla_step_text(580)) ' vla:3077
    End If
    range("c24") = "y" ' vla:3079 src:1008
    vla_step = 581 ' vla:3080
    If vlatraceon() Then ' vla:3081
        Call vlatracestep(581, vla_step_text(581)) ' vla:3081
    End If
    range("a25") = "east" ' vla:3083 src:1009
    vla_step = 582 ' vla:3084
    If vlatraceon() Then ' vla:3085
        Call vlatracestep(582, vla_step_text(582)) ' vla:3085
    End If
    range("b25") = 300 ' vla:3087 src:1010
    vla_step = 583 ' vla:3088
    If vlatraceon() Then ' vla:3089
        Call vlatracestep(583, vla_step_text(583)) ' vla:3089
    End If
    range("c25") = "z" ' vla:3091 src:1011
    vla_step = 584 ' vla:3092
    If vlatraceon() Then ' vla:3093
        Call vlatracestep(584, vla_step_text(584)) ' vla:3093
    End If
    range("a26") = "North" ' vla:3095 src:1012
    vla_step = 585 ' vla:3096
    If vlatraceon() Then ' vla:3097
        Call vlatracestep(585, vla_step_text(585)) ' vla:3097
    End If
    range("b26") = 175 ' vla:3099 src:1013
    vla_step = 586 ' vla:3100
    If vlatraceon() Then ' vla:3101
        Call vlatracestep(586, vla_step_text(586)) ' vla:3101
    End If
    range("c26") = "w" ' vla:3103 src:1014
    vla_step = 587 ' vla:3104
    If vlatraceon() Then ' vla:3105
        Call vlatracestep(587, vla_step_text(587)) ' vla:3105
    End If
    range("a27") = "East" ' vla:3107 src:1015
    vla_step = 588 ' vla:3108
    If vlatraceon() Then ' vla:3109
        Call vlatracestep(588, vla_step_text(588)) ' vla:3109
    End If
    range("b27") = 75 ' vla:3111 src:1016
    vla_step = 589 ' vla:3112
    If vlatraceon() Then ' vla:3113
        Call vlatracestep(589, vla_step_text(589)) ' vla:3113
    End If
    range("c27") = "v" ' vla:3115 src:1017
    vla_step = 590 ' vla:3116
    If vlatraceon() Then ' vla:3117
        Call vlatracestep(590, vla_step_text(590)) ' vla:3117
    End If
    range("b21:b27").numberformat = vlanumberformatcode("dollars", 2)
    vla_step = 591 ' vla:3120
    If vlatraceon() Then ' vla:3121
        Call vlatracestep(591, vla_step_text(591)) ' vla:3121
    End If
    ' ---- instructions.txt:1020 Filter range A20:C27 to show rows where column A is "West". ----
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "a"), criteria1:=vlafiltercriterion("at-least", "West"), operator:=xland, criteria2:=vlafiltercriterion("at-most", "West"))
    vla_step = 592 ' vla:3125
    If vlatraceon() Then ' vla:3126
        Call vlatracestep(592, vla_step_text(592)) ' vla:3126
    End If
    Call range("a20:c27").specialcells(xlcelltypevisible).copy(destination:=range("e30"))
    vla_step = 593 ' vla:3129
    If vlatraceon() Then ' vla:3130
        Call vlatracestep(593, vla_step_text(593)) ' vla:3130
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 594 ' vla:3133
    If vlatraceon() Then ' vla:3134
        Call vlatracestep(594, vla_step_text(594)) ' vla:3134
    End If
    ' ---- instructions.txt:1024 Filter range A20:C27 to show rows where column B is 100. ----
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "b"), criteria1:=vlafiltercriterion("at-least", 100), operator:=xland, criteria2:=vlafiltercriterion("at-most", 100))
    vla_step = 595 ' vla:3138
    If vlatraceon() Then ' vla:3139
        Call vlatracestep(595, vla_step_text(595)) ' vla:3139
    End If
    Call range("a20:c27").specialcells(xlcelltypevisible).copy(destination:=range("i30"))
    vla_step = 596 ' vla:3142
    If vlatraceon() Then ' vla:3143
        Call vlatracestep(596, vla_step_text(596)) ' vla:3143
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 597 ' vla:3146
    If vlatraceon() Then ' vla:3147
        Call vlatracestep(597, vla_step_text(597)) ' vla:3147
    End If
    ' ---- instructions.txt:1028 Filter range A20:C27 to show rows where column C contains "*". ----
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "c"), criteria1:=vlafiltercriterion("contains", "*"))
    vla_step = 598 ' vla:3151
    If vlatraceon() Then ' vla:3152
        Call vlatracestep(598, vla_step_text(598)) ' vla:3152
    End If
    Call range("a20:c27").specialcells(xlcelltypevisible).copy(destination:=range("m30"))
    vla_step = 599 ' vla:3155
    If vlatraceon() Then ' vla:3156
        Call vlatracestep(599, vla_step_text(599)) ' vla:3156
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 600 ' vla:3159
    If vlatraceon() Then ' vla:3160
        Call vlatracestep(600, vla_step_text(600)) ' vla:3160
    End If
    ' ---- instructions.txt:1032 Filter range A20:C27 to show rows where column B is greater than 150. ----
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "b"), criteria1:=vlafiltercriterion("greater", 150))
    vla_step = 601 ' vla:3164
    If vlatraceon() Then ' vla:3165
        Call vlatracestep(601, vla_step_text(601)) ' vla:3165
    End If
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "a"), criteria1:=vlafiltercriterion("at-least", "East"), operator:=xland, criteria2:=vlafiltercriterion("at-most", "East"))
    vla_step = 602 ' vla:3168
    If vlatraceon() Then ' vla:3169
        Call vlatracestep(602, vla_step_text(602)) ' vla:3169
    End If
    Call range("a20:c27").specialcells(xlcelltypevisible).copy(destination:=range("q30"))
    vla_step = 603 ' vla:3172
    If vlatraceon() Then ' vla:3173
        Call vlatracestep(603, vla_step_text(603)) ' vla:3173
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 604 ' vla:3176
    If vlatraceon() Then ' vla:3177
        Call vlatracestep(604, vla_step_text(604)) ' vla:3177
    End If
    ' ---- instructions.txt:1038 Filter range A20:C27 to show rows where column B is less than 99.5. ----
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "b"), criteria1:=vlafiltercriterion("less", 99.5))
    vla_step = 605 ' vla:3181
    If vlatraceon() Then ' vla:3182
        Call vlatracestep(605, vla_step_text(605)) ' vla:3182
    End If
    Call range("a20:c27").specialcells(xlcelltypevisible).copy(destination:=range("u30"))
    vla_step = 606 ' vla:3185
    If vlatraceon() Then ' vla:3186
        Call vlatracestep(606, vla_step_text(606)) ' vla:3186
    End If
    ' ---- instructions.txt:1044 Add filters to range A20:C27. ----
    Call vlaaddfilters(range("a20:c27"))
    vla_step = 607 ' vla:3190
    If vlatraceon() Then ' vla:3191
        Call vlatracestep(607, vla_step_text(607)) ' vla:3191
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 608 ' vla:3194
    If vlatraceon() Then ' vla:3195
        Call vlatracestep(608, vla_step_text(608)) ' vla:3195
    End If
    On Error GoTo vla_tryf_8 ' vla:3197 src:1051
    vla_step = 609 ' vla:3198 src:1051
    If vlatraceon() Then ' vla:3199 src:1051
        Call vlatracestep(609, vla_step_text(609)) ' vla:3199 src:1051
    End If
    Call vlaaddfilters(range("a40:b43"))
    GoTo vla_tryd_8 ' vla:3202 src:1051
vla_tryf_8: ' vla:3203 src:1051
    vla_problem = err.description ' vla:3204 src:1051
    Resume vla_tryr_8 ' vla:3205 src:1051
vla_tryr_8: ' vla:3206 src:1051
    On Error GoTo vla_fail ' vla:3207 src:1051
vla_tryd_8: ' vla:3208 src:1051
    On Error GoTo vla_fail ' vla:3209 src:1051
    vla_step = 610 ' vla:3210
    If vlatraceon() Then ' vla:3211
        Call vlatracestep(610, vla_step_text(610)) ' vla:3211
    End If
    ' ---- instructions.txt:1054 Try: ----
    On Error GoTo vla_tryf_9 ' vla:3214 src:1060
    vla_step = 611 ' vla:3215 src:1060
    If vlatraceon() Then ' vla:3216 src:1060
        Call vlatracestep(611, vla_step_text(611)) ' vla:3216 src:1060
    End If
    application.displayalerts = False
    Call worksheets("gtext").delete
    application.displayalerts = True
    GoTo vla_tryd_9 ' vla:3219 src:1060
vla_tryf_9: ' vla:3220 src:1060
    vla_problem = err.description ' vla:3221 src:1060
    Resume vla_tryr_9 ' vla:3222 src:1060
vla_tryr_9: ' vla:3223 src:1060
    On Error GoTo vla_fail ' vla:3224 src:1060
vla_tryd_9: ' vla:3225 src:1060
    On Error GoTo vla_fail ' vla:3226 src:1060
    vla_step = 612 ' vla:3227
    If vlatraceon() Then ' vla:3228
        Call vlatracestep(612, vla_step_text(612)) ' vla:3228
    End If
    ' ---- instructions.txt:1063 Work on sheet GText. ----
    Call vlaensuresheet("gtext") ' vla:3231 src:1063
    Call worksheets("gtext").activate
    vla_step = 613 ' vla:3232
    If vlatraceon() Then ' vla:3233
        Call vlatracestep(613, vla_step_text(613)) ' vla:3233
    End If
    ' ---- instructions.txt:1065 Put "widget" into cell A1. ----
    range("a1") = "widget" ' vla:3236 src:1067
    vla_step = 614 ' vla:3237
    If vlatraceon() Then ' vla:3238
        Call vlatracestep(614, vla_step_text(614)) ' vla:3238
    End If
    range("a2") = 42 ' vla:3240 src:1068
    vla_step = 615 ' vla:3241
    If vlatraceon() Then ' vla:3242
        Call vlatracestep(615, vla_step_text(615)) ' vla:3242
    End If
    range("a3").Formula2 = "=CHAR(97)&CHAR(98)"
    vla_step = 616 ' vla:3245
    If vlatraceon() Then ' vla:3246
        Call vlatracestep(616, vla_step_text(616)) ' vla:3246
    End If
    range("a4") = "'true" ' vla:3248 src:1070
    vla_step = 617 ' vla:3249
    If vlatraceon() Then ' vla:3250
        Call vlatracestep(617, vla_step_text(617)) ' vla:3250
    End If
    range("a5") = "'=abc" ' vla:3252 src:1071
    vla_step = 618 ' vla:3253
    If vlatraceon() Then ' vla:3254
        Call vlatracestep(618, vla_step_text(618)) ' vla:3254
    End If
    Call vlatextinrange(range("a1:a5"), "upper")
    vla_step = 619 ' vla:3257
    If vlatraceon() Then ' vla:3258
        Call vlatracestep(619, vla_step_text(619)) ' vla:3258
    End If
    ' ---- instructions.txt:1074 Put "HELLO World" into cell B1. ----
    range("b1") = "HELLO World" ' vla:3261 src:1075
    vla_step = 620 ' vla:3262
    If vlatraceon() Then ' vla:3263
        Call vlatracestep(620, vla_step_text(620)) ' vla:3263
    End If
    range("b2") = "MiXeD" ' vla:3265 src:1076
    vla_step = 621 ' vla:3266
    If vlatraceon() Then ' vla:3267
        Call vlatracestep(621, vla_step_text(621)) ' vla:3267
    End If
    Call vlatextinrange(columns("b"), "lower")
    vla_step = 622 ' vla:3270
    If vlatraceon() Then ' vla:3271
        Call vlatracestep(622, vla_step_text(622)) ' vla:3271
    End If
    ' ---- instructions.txt:1079 Put "don't stop" into cell C1. ----
    range("c1") = "don't stop" ' vla:3274 src:1080
    vla_step = 623 ' vla:3275
    If vlatraceon() Then ' vla:3276
        Call vlatracestep(623, vla_step_text(623)) ' vla:3276
    End If
    range("c2") = "3rd quarter" ' vla:3278 src:1081
    vla_step = 624 ' vla:3279
    If vlatraceon() Then ' vla:3280
        Call vlatracestep(624, vla_step_text(624)) ' vla:3280
    End If
    range("c3") = "o'neil" ' vla:3282 src:1082
    vla_step = 625 ' vla:3283
    If vlatraceon() Then ' vla:3284
        Call vlatracestep(625, vla_step_text(625)) ' vla:3284
    End If
    Call vlatextinrange(range("c1:c3"), "capitalize-after-space")
    vla_step = 626 ' vla:3287
    If vlatraceon() Then ' vla:3288
        Call vlatracestep(626, vla_step_text(626)) ' vla:3288
    End If
    range("d1") = "don't stop" ' vla:3290 src:1084
    vla_step = 627 ' vla:3291
    If vlatraceon() Then ' vla:3292
        Call vlatracestep(627, vla_step_text(627)) ' vla:3292
    End If
    range("d2") = "3rd quarter" ' vla:3294 src:1085
    vla_step = 628 ' vla:3295
    If vlatraceon() Then ' vla:3296
        Call vlatracestep(628, vla_step_text(628)) ' vla:3296
    End If
    range("d3") = "o'neil" ' vla:3298 src:1086
    vla_step = 629 ' vla:3299
    If vlatraceon() Then ' vla:3300
        Call vlatracestep(629, vla_step_text(629)) ' vla:3300
    End If
    Call vlatextinrange(columns("d"), "capitalize-after-non-letter")
    vla_step = 630 ' vla:3303
    If vlatraceon() Then ' vla:3304
        Call vlatracestep(630, vla_step_text(630)) ' vla:3304
    End If
    ' ---- instructions.txt:1089 Put "  a   b  " into cell E1. ----
    range("e1") = "  a   b  " ' vla:3307 src:1092
    vla_step = 631 ' vla:3308
    If vlatraceon() Then ' vla:3309
        Call vlatracestep(631, vla_step_text(631)) ' vla:3309
    End If
    range("e2") = "'  00123 " ' vla:3311 src:1093
    vla_step = 632 ' vla:3312
    If vlatraceon() Then ' vla:3313
        Call vlatracestep(632, vla_step_text(632)) ' vla:3313
    End If
    range("e3") = "   " ' vla:3315 src:1094
    vla_step = 633 ' vla:3316
    If vlatraceon() Then ' vla:3317
        Call vlatracestep(633, vla_step_text(633)) ' vla:3317
    End If
    range("e4").Formula2 = "=UNICHAR(160)&CHAR(120)&UNICHAR(160)&UNICHAR(160)&CHAR(121)"
    vla_step = 634 ' vla:3320
    If vlatraceon() Then ' vla:3321
        Call vlatracestep(634, vla_step_text(634)) ' vla:3321
    End If
    Call range("e4:e4").copy
    Call range("e4:e4").pastespecial(paste:=xlpastevalues)
    application.cutcopymode = False
    vla_step = 635 ' vla:3324
    If vlatraceon() Then ' vla:3325
        Call vlatracestep(635, vla_step_text(635)) ' vla:3325
    End If
    Call vlatextinrange(range("e1:e4"), "remove-extra-spaces")
    vla_step = 636 ' vla:3328
    If vlatraceon() Then ' vla:3329
        Call vlatracestep(636, vla_step_text(636)) ' vla:3329
    End If
    ' ---- instructions.txt:1099 Put formula "=CHAR(97)&CHAR(10)&CHAR(98)" into cell F1. ----
    range("f1").Formula2 = "=CHAR(97)&CHAR(10)&CHAR(98)"
    vla_step = 637 ' vla:3333
    If vlatraceon() Then ' vla:3334
        Call vlatracestep(637, vla_step_text(637)) ' vla:3334
    End If
    range("f2").Formula2 = "=CHAR(99)&CHAR(9)&CHAR(100)"
    vla_step = 638 ' vla:3337
    If vlatraceon() Then ' vla:3338
        Call vlatracestep(638, vla_step_text(638)) ' vla:3338
    End If
    Call range("f1:f2").copy
    Call range("f1:f2").pastespecial(paste:=xlpastevalues)
    application.cutcopymode = False
    vla_step = 639 ' vla:3341
    If vlatraceon() Then ' vla:3342
        Call vlatracestep(639, vla_step_text(639)) ' vla:3342
    End If
    Call vlatextinrange(range("f1:f2"), "remove-non-printing")
    vla_step = 640 ' vla:3345
    If vlatraceon() Then ' vla:3346
        Call vlatracestep(640, vla_step_text(640)) ' vla:3346
    End If
    ' ---- instructions.txt:1105 Put "Widget" into cell G1. ----
    range("g1") = "Widget" ' vla:3349 src:1108
    vla_step = 641 ' vla:3350
    If vlatraceon() Then ' vla:3351
        Call vlatracestep(641, vla_step_text(641)) ' vla:3351
    End If
    range("g2") = "banana" ' vla:3353 src:1109
    vla_step = 642 ' vla:3354
    If vlatraceon() Then ' vla:3355
        Call vlatracestep(642, vla_step_text(642)) ' vla:3355
    End If
    gtext_row = vlafindrow("Widget", "g") ' vla:3357 src:1110
    vla_step = 643 ' vla:3358
    If vlatraceon() Then ' vla:3359
        Call vlatracestep(643, vla_step_text(643)) ' vla:3359
    End If
    Call vlareplaceinrange(range("g2:g2"), "AN", "xy", "values")
    vla_step = 644 ' vla:3362
    If vlatraceon() Then ' vla:3363
        Call vlatracestep(644, vla_step_text(644)) ' vla:3363
    End If
    ' ---- instructions.txt:1113 Format range I1:I20 as text for new entries. ----
    range("i1:i20").numberformat = "@"
    vla_step = 645 ' vla:3367
    If vlatraceon() Then ' vla:3368
        Call vlatracestep(645, vla_step_text(645)) ' vla:3368
    End If
    gtext_code = "INV-ab-Cd" ' vla:3370 src:1116
    vla_step = 646 ' vla:3371
    If vlatraceon() Then ' vla:3372
        Call vlatracestep(646, vla_step_text(646)) ' vla:3372
    End If
    gtext_part = vlatextbeside(gtext_code, "-", "before", "first") ' vla:3374 src:1117
    vla_step = 647 ' vla:3375
    If vlatraceon() Then ' vla:3376
        Call vlatracestep(647, vla_step_text(647)) ' vla:3376
    End If
    range("i1") = gtext_part ' vla:3378 src:1118
    vla_step = 648 ' vla:3379
    If vlatraceon() Then ' vla:3380
        Call vlatracestep(648, vla_step_text(648)) ' vla:3380
    End If
    gtext_part = vlatextbeside(gtext_code, "-", "after", "first") ' vla:3382 src:1119
    vla_step = 649 ' vla:3383
    If vlatraceon() Then ' vla:3384
        Call vlatracestep(649, vla_step_text(649)) ' vla:3384
    End If
    range("i2") = gtext_part ' vla:3386 src:1120
    vla_step = 650 ' vla:3387
    If vlatraceon() Then ' vla:3388
        Call vlatracestep(650, vla_step_text(650)) ' vla:3388
    End If
    gtext_part = vlatextbeside(gtext_code, "-", "after", "last") ' vla:3390 src:1121
    vla_step = 651 ' vla:3391
    If vlatraceon() Then ' vla:3392
        Call vlatracestep(651, vla_step_text(651)) ' vla:3392
    End If
    range("i3") = gtext_part ' vla:3394 src:1122
    vla_step = 652 ' vla:3395
    If vlatraceon() Then ' vla:3396
        Call vlatracestep(652, vla_step_text(652)) ' vla:3396
    End If
    ' ---- instructions.txt:1124 Set gtext-part to the text before "c" in gtext-code. ----
    gtext_part = vlatextbeside(gtext_code, "c", "before", "first") ' vla:3399 src:1125
    vla_step = 653 ' vla:3400
    If vlatraceon() Then ' vla:3401
        Call vlatracestep(653, vla_step_text(653)) ' vla:3401
    End If
    range("i4") = gtext_part ' vla:3403 src:1126
    vla_step = 654 ' vla:3404
    If vlatraceon() Then ' vla:3405
        Call vlatracestep(654, vla_step_text(654)) ' vla:3405
    End If
    gtext_part = left(gtext_code, 3) ' vla:3407 src:1127
    vla_step = 655 ' vla:3408
    If vlatraceon() Then ' vla:3409
        Call vlatracestep(655, vla_step_text(655)) ' vla:3409
    End If
    range("i5") = gtext_part ' vla:3411 src:1128
    vla_step = 656 ' vla:3412
    If vlatraceon() Then ' vla:3413
        Call vlatracestep(656, vla_step_text(656)) ' vla:3413
    End If
    gtext_part = right(gtext_code, 2) ' vla:3415 src:1129
    vla_step = 657 ' vla:3416
    If vlatraceon() Then ' vla:3417
        Call vlatracestep(657, vla_step_text(657)) ' vla:3417
    End If
    range("i6") = gtext_part ' vla:3419 src:1130
    vla_step = 658 ' vla:3420
    If vlatraceon() Then ' vla:3421
        Call vlatracestep(658, vla_step_text(658)) ' vla:3421
    End If
    ' ---- instructions.txt:1132 Set gtext-part to 42 padded on the left with "0" to 5 characters. ----
    gtext_part = vlatextpad(42, "0", 5, "left") ' vla:3424 src:1132
    vla_step = 659 ' vla:3425
    If vlatraceon() Then ' vla:3426
        Call vlatracestep(659, vla_step_text(659)) ' vla:3426
    End If
    range("i7") = gtext_part ' vla:3428 src:1133
    vla_step = 660 ' vla:3429
    If vlatraceon() Then ' vla:3430
        Call vlatracestep(660, vla_step_text(660)) ' vla:3430
    End If
    gtext_part = vlatextpad("ab", ".", 4, "right") ' vla:3432 src:1134
    vla_step = 661 ' vla:3433
    If vlatraceon() Then ' vla:3434
        Call vlatracestep(661, vla_step_text(661)) ' vla:3434
    End If
    range("i8") = gtext_part ' vla:3436 src:1135
    vla_step = 662 ' vla:3437
    If vlatraceon() Then ' vla:3438
        Call vlatracestep(662, vla_step_text(662)) ' vla:3438
    End If
    ' ---- instructions.txt:1137 Put "x" into cell J1. ----
    range("j1") = "x" ' vla:3441 src:1138
    vla_step = 663 ' vla:3442
    If vlatraceon() Then ' vla:3443
        Call vlatracestep(663, vla_step_text(663)) ' vla:3443
    End If
    range("j3") = 7 ' vla:3445 src:1139
    vla_step = 664 ' vla:3446
    If vlatraceon() Then ' vla:3447
        Call vlatracestep(664, vla_step_text(664)) ' vla:3447
    End If
    range("j4") = "y" ' vla:3449 src:1140
    vla_step = 665 ' vla:3450
    If vlatraceon() Then ' vla:3451
        Call vlatracestep(665, vla_step_text(665)) ' vla:3451
    End If
    gtext_list = vlajoinrange(range("j1:j4"), ", ") ' vla:3453 src:1141
    vla_step = 666 ' vla:3454
    If vlatraceon() Then ' vla:3455
        Call vlatracestep(666, vla_step_text(666)) ' vla:3455
    End If
    range("i9") = gtext_list ' vla:3457 src:1142
    vla_step = 667 ' vla:3458
    If vlatraceon() Then ' vla:3459
        Call vlatracestep(667, vla_step_text(667)) ' vla:3459
    End If
    gtext_list = vlajoinrange(range("j1:j4"), "; ") ' vla:3461 src:1143
    vla_step = 668 ' vla:3462
    If vlatraceon() Then ' vla:3463
        Call vlatracestep(668, vla_step_text(668)) ' vla:3463
    End If
    range("i10") = gtext_list ' vla:3465 src:1144
    vla_step = 669 ' vla:3466
    If vlatraceon() Then ' vla:3467
        Call vlatracestep(669, vla_step_text(669)) ' vla:3467
    End If
    gtext_list = vlajoinrange(columns("j"), ", ") ' vla:3469 src:1145
    vla_step = 670 ' vla:3470
    If vlatraceon() Then ' vla:3471
        Call vlatracestep(670, vla_step_text(670)) ' vla:3471
    End If
    range("i11") = gtext_list ' vla:3473 src:1146
    vla_step = 671 ' vla:3474
    If vlatraceon() Then ' vla:3475
        Call vlatracestep(671, vla_step_text(671)) ' vla:3475
    End If
    ' ---- instructions.txt:1148 Set gtext-part to "  a   b " with extra spaces removed. ----
    gtext_part = vlatextop("  a   b ", "remove-extra-spaces") ' vla:3478 src:1150
    vla_step = 672 ' vla:3479
    If vlatraceon() Then ' vla:3480
        Call vlatracestep(672, vla_step_text(672)) ' vla:3480
    End If
    range("i12") = gtext_part ' vla:3482 src:1151
    vla_step = 673 ' vla:3483
    If vlatraceon() Then ' vla:3484
        Call vlatracestep(673, vla_step_text(673)) ' vla:3484
    End If
    gtext_part = vlatextop("o'neil", "capitalize-after-non-letter") ' vla:3486 src:1152
    vla_step = 674 ' vla:3487
    If vlatraceon() Then ' vla:3488
        Call vlatracestep(674, vla_step_text(674)) ' vla:3488
    End If
    range("i13") = gtext_part ' vla:3490 src:1153
    vla_step = 675 ' vla:3491
    If vlatraceon() Then ' vla:3492
        Call vlatracestep(675, vla_step_text(675)) ' vla:3492
    End If
    gtext_part = vlatextop("don't stop", "capitalize-after-space") ' vla:3494 src:1154
    vla_step = 676 ' vla:3495
    If vlatraceon() Then ' vla:3496
        Call vlatracestep(676, vla_step_text(676)) ' vla:3496
    End If
    range("i14") = gtext_part ' vla:3498 src:1155
    vla_step = 677 ' vla:3499
    If vlatraceon() Then ' vla:3500
        Call vlatracestep(677, vla_step_text(677)) ' vla:3500
    End If
    range("k1").Formula2 = "=CHAR(112)&CHAR(10)&CHAR(113)"
    vla_step = 678 ' vla:3503
    If vlatraceon() Then ' vla:3504
        Call vlatracestep(678, vla_step_text(678)) ' vla:3504
    End If
    Call range("k1:k1").copy
    Call range("k1:k1").pastespecial(paste:=xlpastevalues)
    application.cutcopymode = False
    vla_step = 679 ' vla:3507
    If vlatraceon() Then ' vla:3508
        Call vlatracestep(679, vla_step_text(679)) ' vla:3508
    End If
    gtext_part = vlatextop(range("k1"), "remove-non-printing") ' vla:3510 src:1158
    vla_step = 680 ' vla:3511
    If vlatraceon() Then ' vla:3512
        Call vlatracestep(680, vla_step_text(680)) ' vla:3512
    End If
    range("i15") = gtext_part ' vla:3514 src:1159
    vla_step = 681 ' vla:3515
    If vlatraceon() Then ' vla:3516
        Call vlatracestep(681, vla_step_text(681)) ' vla:3516
    End If
    ' ---- instructions.txt:1161 Try: ----
    On Error GoTo vla_tryf_10 ' vla:3519 src:1164
    vla_step = 682 ' vla:3520 src:1164
    If vlatraceon() Then ' vla:3521 src:1164
        Call vlatracestep(682, vla_step_text(682)) ' vla:3521 src:1164
    End If
    gtext_part = vlatextbeside(gtext_code, "#", "before", "first") ' vla:3523 src:1165
    GoTo vla_tryd_10 ' vla:3524 src:1164
vla_tryf_10: ' vla:3525 src:1164
    vla_problem = err.description ' vla:3526 src:1164
    Resume vla_tryr_10 ' vla:3527 src:1164
vla_tryr_10: ' vla:3528 src:1164
    On Error GoTo vla_fail ' vla:3529 src:1164
vla_tryd_10: ' vla:3530 src:1164
    On Error GoTo vla_fail ' vla:3531 src:1164
    vla_step = 683 ' vla:3532
    If vlatraceon() Then ' vla:3533
        Call vlatracestep(683, vla_step_text(683)) ' vla:3533
    End If
    ' ---- instructions.txt:1167 Put gtext-part into cell I16. ----
    range("i16") = gtext_part ' vla:3536 src:1167
    vla_step = 684 ' vla:3537
    If vlatraceon() Then ' vla:3538
        Call vlatracestep(684, vla_step_text(684)) ' vla:3538
    End If
    ' ---- instructions.txt:1169 Put "Nan" into cell L1. ----
    range("l1") = "Nan" ' vla:3541 src:1173
    vla_step = 685 ' vla:3542
    If vlatraceon() Then ' vla:3543
        Call vlatracestep(685, vla_step_text(685)) ' vla:3543
    End If
    range("l2").Formula2 = "=TAN(0)"
    vla_step = 686 ' vla:3546
    If vlatraceon() Then ' vla:3547
        Call vlatracestep(686, vla_step_text(686)) ' vla:3547
    End If
    Call vlareplaceinrange(range("l1:l2"), "an", "xy", "values")
    vla_step = 687 ' vla:3550
    If vlatraceon() Then ' vla:3551
        Call vlatracestep(687, vla_step_text(687)) ' vla:3551
    End If
    range("l3") = "Q" ' vla:3553 src:1176
    vla_step = 688 ' vla:3554
    If vlatraceon() Then ' vla:3555
        Call vlatracestep(688, vla_step_text(688)) ' vla:3555
    End If
    Call vlareplaceinrange(range("l3:l3"), "Q", "=1+1", "values")
    vla_step = 689 ' vla:3558
    If vlatraceon() Then ' vla:3559
        Call vlatracestep(689, vla_step_text(689)) ' vla:3559
    End If
    range("l4") = "N/A" ' vla:3561 src:1178
    vla_step = 690 ' vla:3562
    If vlatraceon() Then ' vla:3563
        Call vlatracestep(690, vla_step_text(690)) ' vla:3563
    End If
    Call vlareplaceinrange(range("l4:l4"), "N/A", 0, "values")
    vla_step = 691 ' vla:3566
    If vlatraceon() Then ' vla:3567
        Call vlatracestep(691, vla_step_text(691)) ' vla:3567
    End If
    ' ---- instructions.txt:1181 Put formula "=ABS(-2)" into cell L5. ----
    range("l5").Formula2 = "=ABS(-2)"
    vla_step = 692 ' vla:3571
    If vlatraceon() Then ' vla:3572
        Call vlatracestep(692, vla_step_text(692)) ' vla:3572
    End If
    range("l6") = "ABS" ' vla:3574 src:1183
    vla_step = 693 ' vla:3575
    If vlatraceon() Then ' vla:3576
        Call vlatracestep(693, vla_step_text(693)) ' vla:3576
    End If
    Call vlareplaceinrange(range("l5:l6"), "ABS", "SIGN", "formulas")
    vla_step = 694 ' vla:3579
    If vlatraceon() Then ' vla:3580
        Call vlatracestep(694, vla_step_text(694)) ' vla:3580
    End If
    ' ---- instructions.txt:1186 Put "zqz" into cell L7. ----
    range("l7") = "zqz" ' vla:3583 src:1187
    vla_step = 695 ' vla:3584
    If vlatraceon() Then ' vla:3585
        Call vlatracestep(695, vla_step_text(695)) ' vla:3585
    End If
    Call vlareplaceinrange(activesheet.usedrange, "zqz", "done", "values")
    vla_step = 696 ' vla:3588
    If vlatraceon() Then ' vla:3589
        Call vlatracestep(696, vla_step_text(696)) ' vla:3589
    End If
    range("l8").Formula2 = "=ABS(-3)"
    vla_step = 697 ' vla:3592
    If vlatraceon() Then ' vla:3593
        Call vlatracestep(697, vla_step_text(697)) ' vla:3593
    End If
    Call vlareplaceinrange(activesheet.usedrange, "-3", "-4", "formulas")
    vla_step = 698 ' vla:3596
    If vlatraceon() Then ' vla:3597
        Call vlatracestep(698, vla_step_text(698)) ' vla:3597
    End If
    ' ---- instructions.txt:1192 Put "x,0042,7" into cell N1. ----
    range("n1") = "x,0042,7" ' vla:3600 src:1195
    vla_step = 699 ' vla:3601
    If vlatraceon() Then ' vla:3602
        Call vlatracestep(699, vla_step_text(699)) ' vla:3602
    End If
    Call vlasplitcolumn(columns("n"), ",", "text")
    vla_step = 700 ' vla:3605
    If vlatraceon() Then ' vla:3606
        Call vlatracestep(700, vla_step_text(700)) ' vla:3606
    End If
    range("r1") = "y,0042,7,-3.5" ' vla:3608 src:1197
    vla_step = 701 ' vla:3609
    If vlatraceon() Then ' vla:3610
        Call vlatracestep(701, vla_step_text(701)) ' vla:3610
    End If
    Call vlasplitcolumn(columns("r"), ",", "numbers")
    vla_step = 702 ' vla:3613
    If vlatraceon() Then ' vla:3614
        Call vlatracestep(702, vla_step_text(702)) ' vla:3614
    End If
    range("w1") = "p,q" ' vla:3616 src:1199
    vla_step = 703 ' vla:3617
    If vlatraceon() Then ' vla:3618
        Call vlatracestep(703, vla_step_text(703)) ' vla:3618
    End If
    range("x1") = "keep" ' vla:3620 src:1200
    vla_step = 704 ' vla:3621
    If vlatraceon() Then ' vla:3622
        Call vlatracestep(704, vla_step_text(704)) ' vla:3622
    End If
    On Error GoTo vla_tryf_11 ' vla:3624 src:1201
    vla_step = 705 ' vla:3625 src:1201
    If vlatraceon() Then ' vla:3626 src:1201
        Call vlatracestep(705, vla_step_text(705)) ' vla:3626 src:1201
    End If
    Call vlasplitcolumn(columns("w"), ",", "text")
    GoTo vla_tryd_11 ' vla:3629 src:1201
vla_tryf_11: ' vla:3630 src:1201
    vla_problem = err.description ' vla:3631 src:1201
    Resume vla_tryr_11 ' vla:3632 src:1201
vla_tryr_11: ' vla:3633 src:1201
    On Error GoTo vla_fail ' vla:3634 src:1201
vla_tryd_11: ' vla:3635 src:1201
    On Error GoTo vla_fail ' vla:3636 src:1201
    vla_step = 706 ' vla:3637
    If vlatraceon() Then ' vla:3638
        Call vlatracestep(706, vla_step_text(706)) ' vla:3638
    End If
    ' ---- instructions.txt:1204 Put "apple" into cell Z1. ----
    range("z1") = "apple" ' vla:3641 src:1205
    vla_step = 707 ' vla:3642
    If vlatraceon() Then ' vla:3643
        Call vlatracestep(707, vla_step_text(707)) ' vla:3643
    End If
    range("aa1") = "Banana" ' vla:3645 src:1206
    vla_step = 708 ' vla:3646
    If vlatraceon() Then ' vla:3647
        Call vlatracestep(708, vla_step_text(708)) ' vla:3647
    End If
    range("z2") = "cherry" ' vla:3649 src:1207
    vla_step = 709 ' vla:3650
    If vlatraceon() Then ' vla:3651
        Call vlatracestep(709, vla_step_text(709)) ' vla:3651
    End If
    range("aa2") = "banana split" ' vla:3653 src:1208
    vla_step = 710 ' vla:3654
    If vlatraceon() Then ' vla:3655
        Call vlatracestep(710, vla_step_text(710)) ' vla:3655
    End If
    range("z3") = 7 ' vla:3657 src:1209
    vla_step = 711 ' vla:3658
    If vlatraceon() Then ' vla:3659
        Call vlatracestep(711, vla_step_text(711)) ' vla:3659
    End If
    range("aa3") = "BANANA" ' vla:3661 src:1210
    vla_step = 712 ' vla:3662
    If vlatraceon() Then ' vla:3663
        Call vlatracestep(712, vla_step_text(712)) ' vla:3663
    End If
    gtext_found = vlafindtext(range("z1:aa3"), "nan", "row") ' vla:3665 src:1211
    vla_step = 713 ' vla:3666
    If vlatraceon() Then ' vla:3667
        Call vlatracestep(713, vla_step_text(713)) ' vla:3667
    End If
    range("l9") = gtext_found ' vla:3669 src:1212
    vla_step = 714 ' vla:3670
    If vlatraceon() Then ' vla:3671
        Call vlatracestep(714, vla_step_text(714)) ' vla:3671
    End If
    gtext_found = vlafindtext(range("z1:aa3"), "nan", "column") ' vla:3673 src:1213
    vla_step = 715 ' vla:3674
    If vlatraceon() Then ' vla:3675
        Call vlatracestep(715, vla_step_text(715)) ' vla:3675
    End If
    range("l10") = gtext_found ' vla:3677 src:1214
    vla_step = 716 ' vla:3678
    If vlatraceon() Then ' vla:3679
        Call vlatracestep(716, vla_step_text(716)) ' vla:3679
    End If
    gtext_found = vlafindtext(columns("z"), "err", "row") ' vla:3681 src:1215
    vla_step = 717 ' vla:3682
    If vlatraceon() Then ' vla:3683
        Call vlatracestep(717, vla_step_text(717)) ' vla:3683
    End If
    range("l11") = gtext_found ' vla:3685 src:1216
    vla_step = 718 ' vla:3686
    If vlatraceon() Then ' vla:3687
        Call vlatracestep(718, vla_step_text(718)) ' vla:3687
    End If
    gtext_found = vlacounttext(range("z1:aa3"), "banana") ' vla:3689 src:1217
    vla_step = 719 ' vla:3690
    If vlatraceon() Then ' vla:3691
        Call vlatracestep(719, vla_step_text(719)) ' vla:3691
    End If
    range("l12") = gtext_found ' vla:3693 src:1218
    vla_step = 720 ' vla:3694
    If vlatraceon() Then ' vla:3695
        Call vlatracestep(720, vla_step_text(720)) ' vla:3695
    End If
    gtext_found = vlacounttext(columns("z"), "zzz") ' vla:3697 src:1219
    vla_step = 721 ' vla:3698
    If vlatraceon() Then ' vla:3699
        Call vlatracestep(721, vla_step_text(721)) ' vla:3699
    End If
    range("l13") = gtext_found ' vla:3701 src:1220
    vla_step = 722 ' vla:3702
    If vlatraceon() Then ' vla:3703
        Call vlatracestep(722, vla_step_text(722)) ' vla:3703
    End If
    gtext_found = vlafindtext(range("z1:aa3"), "kiwi", "row") ' vla:3705 src:1221
    vla_step = 723 ' vla:3706
    If vlatraceon() Then ' vla:3707
        Call vlatracestep(723, vla_step_text(723)) ' vla:3707
    End If
    range("l14") = gtext_found ' vla:3709 src:1222
    vla_step = 724 ' vla:3710
    If vlatraceon() Then ' vla:3711
        Call vlatracestep(724, vla_step_text(724)) ' vla:3711
    End If
    If (vlafindtext(range("z1:aa3"), "split", "row") > 0) Then ' vla:3713 src:1223
        range("l15") = "yes" ' vla:3713 src:1223
    End If
    vla_step = 725 ' vla:3714
    If vlatraceon() Then ' vla:3715
        Call vlatracestep(725, vla_step_text(725)) ' vla:3715
    End If
    If (vlafindtext(columns("z"), "kiwi", "row") = 0) Then ' vla:3717 src:1224
        range("l16") = "no kiwi" ' vla:3717 src:1224
    End If
    vla_step = 735 ' vla:3718
    If vlatraceon() Then ' vla:3719
        Call vlatracestep(735, vla_step_text(735)) ' vla:3719
    End If
    ' ---- instructions.txt:1248 Check-text-conditions. ----
    Call check_text_conditions ' vla:3722 src:1248
    vla_step = 761 ' vla:3723
    If vlatraceon() Then ' vla:3724
        Call vlatracestep(761, vla_step_text(761)) ' vla:3724
    End If
    ' ---- instructions.txt:1285 Check-formulas. ----
    Call check_formulas ' vla:3727 src:1285
    vla_step = 789 ' vla:3728
    If vlatraceon() Then ' vla:3729
        Call vlatracestep(789, vla_step_text(789)) ' vla:3729
    End If
    ' ---- instructions.txt:1330 Check-statistics. ----
    Call check_statistics ' vla:3732 src:1330
    vla_step = 812 ' vla:3733
    If vlatraceon() Then ' vla:3734
        Call vlatracestep(812, vla_step_text(812)) ' vla:3734
    End If
    ' ---- instructions.txt:1366 Check-calculation. ----
    Call check_calculation ' vla:3737 src:1366
    vla_step = 842 ' vla:3738
    If vlatraceon() Then ' vla:3739
        Call vlatracestep(842, vla_step_text(842)) ' vla:3739
    End If
    ' ---- instructions.txt:1410 Check-figures. ----
    Call check_figures ' vla:3742 src:1410
    vla_step = 843 ' vla:3743
    If vlatraceon() Then ' vla:3744
        Call vlatracestep(843, vla_step_text(843)) ' vla:3744
    End If
    ' ---- instructions.txt:1412 Go to sheet Output. ----
    Call worksheets("output").activate
    vla_step = 844 ' vla:3748
    If vlatraceon() Then ' vla:3749
        Call vlatracestep(844, vla_step_text(844)) ' vla:3749
    End If
    ' ---- instructions.txt:1414 Tidy-up. ----
    Call tidy_up ' vla:3752 src:1414
    vla_step = 845 ' vla:3753
    If vlatraceon() Then ' vla:3754
        Call vlatracestep(845, vla_step_text(845)) ' vla:3754
    End If
    application.screenupdating = True
    vla_step = 846 ' vla:3757
    If vlatraceon() Then ' vla:3758
        Call vlatracestep(846, vla_step_text(846)) ' vla:3758
    End If
    Debug.Print "report finished" ' vla:3760 src:1416
    Exit Sub ' vla:3761
vla_fail: ' vla:3762
    Call vla_report_error ' vla:3763
End Sub

Public Sub vla_report_error()
    Call vlareportstop(vla_step, vla_step_text(vla_step), err.description) ' vla:3766
End Sub

Public Function vla_step_text(ByVal n As Long) As String
    Select Case n ' vla:3769
        Case 1
            vla_step_text = "Work on sheet Output. [line 18]" ' vla:3770
            Exit Function
        Case 2
            vla_step_text = "Turn off screen updating. [line 20]" ' vla:3771
            Exit Function
        Case 3
            vla_step_text = "Create a number called total. [line 21]" ' vla:3772
            Exit Function
        Case 4
            vla_step_text = "Set total to 0. [line 22]" ' vla:3773
            Exit Function
        Case 5
            vla_step_text = "Put ""Test Report"" into cell A1. [line 23]" ' vla:3774
            Exit Function
        Case 6
            vla_step_text = "Make cell A1 bold. [line 24]" ' vla:3775
            Exit Function
        Case 7
            vla_step_text = "Set font size of cell A1 to 14. [line 25]" ' vla:3776
            Exit Function
        Case 8
            vla_step_text = "Put today into cell D1. [line 26]" ' vla:3777
            Exit Function
        Case 9
            vla_step_text = "Repeat 5 times: [line 28]" ' vla:3778
            Exit Function
        Case 10
            vla_step_text = "Increase total by counter. [line 29]" ' vla:3779
            Exit Function
        Case 11
            vla_step_text = "If counter is divisible by 2, log ""even step "" joined with counter. [line 30]" ' vla:3780
            Exit Function
        Case 12
            vla_step_text = "Put total into cell B2. [line 32]" ' vla:3781
            Exit Function
        Case 13
            vla_step_text = "Put formula ""=B2*2"" into cell B3. [line 33]" ' vla:3782
            Exit Function
        Case 14
            vla_step_text = "Put sum of range ""B2:B3"" into cell B4. [line 37]" ' vla:3783
            Exit Function
        Case 15
            vla_step_text = "Set grand to sum of range B2:B3. [line 38]" ' vla:3784
            Exit Function
        Case 16
            vla_step_text = "Log ""grand is "" joined with grand. [line 39]" ' vla:3785
            Exit Function
        Case 17
            vla_step_text = "If grand is greater than 40: [line 43]" ' vla:3786
            Exit Function
        Case 18
            vla_step_text = "Put ""PASS"" into cell C4. [line 44]" ' vla:3787
            Exit Function
        Case 19
            vla_step_text = "Repeat 2 times: [line 45]" ' vla:3788
            Exit Function
        Case 20
            vla_step_text = "Log ""pass check "" joined with counter. [line 46]" ' vla:3789
            Exit Function
        Case 21
            vla_step_text = "Make cell C4 bold. [line 48]" ' vla:3790
            Exit Function
        Case 22
            vla_step_text = "Put ""CHECK"" into cell C4. [line 51]" ' vla:3791
            Exit Function
        Case 23
            vla_step_text = "If grand is at least 45, make cell C4 yellow. [line 53]" ' vla:3792
            Exit Function
        Case 24
            vla_step_text = "Create a text called label. [line 54]" ' vla:3793
            Exit Function
        Case 25
            vla_step_text = "Set label to ""Total: "" joined with total. [line 55]" ' vla:3794
            Exit Function
        Case 26
            vla_step_text = "Put label into cell A6. [line 56]" ' vla:3795
            Exit Function
        Case 27
            vla_step_text = "Set font-color of cell A6 to hot-pink. [line 57]" ' vla:3796
            Exit Function
        Case 28
            vla_step_text = "Remember range B2:B4 as results. [line 58]" ' vla:3797
            Exit Function
        Case 29
            vla_step_text = "For each r in results, log r. [line 59]" ' vla:3798
            Exit Function
        Case 30
            vla_step_text = "Set biggest to largest of results. [line 60]" ' vla:3799
            Exit Function
        Case 31
            vla_step_text = "Log ""largest result is "" joined with biggest joined with "", label length "" joined with length of label. [line 61]" ' vla:3800
            Exit Function
        Case 32
            vla_step_text = "Repeat 3 times: [line 64]" ' vla:3801
            Exit Function
        Case 33
            vla_step_text = "Put counter times 10 into column E row counter. [line 65]" ' vla:3802
            Exit Function
        Case 34
            vla_step_text = "Set probe to cell in column E row 2. [line 67]" ' vla:3803
            Exit Function
        Case 35
            vla_step_text = "Create a text called num-col-check. [line 68]" ' vla:3804
            Exit Function
        Case 36
            vla_step_text = "If cell in column number 5 row 2 is 20, set num-col-check to ""yes"". [line 69]" ' vla:3805
            Exit Function
        Case 37
            vla_step_text = "Create a text called value-word-check. [line 70]" ' vla:3806
            Exit Function
        Case 38
            vla_step_text = "If value in column number 5 row 3 is 30, set value-word-check to ""ok"". [line 71]" ' vla:3807
            Exit Function
        Case 39
            vla_step_text = "Put ""bang"" into cell Output!H16. [line 72]" ' vla:3808
            Exit Function
        Case 40
            vla_step_text = "Center cell A1. [line 73]" ' vla:3809
            Exit Function
        Case 41
            vla_step_text = "Set width of column D to 24. [line 74]" ' vla:3810
            Exit Function
        Case 42
            vla_step_text = "Set round-check to 3.14159 rounded to 2 decimals. [line 75]" ' vla:3811
            Exit Function
        Case 43
            vla_step_text = "Log ""rounded is "" joined with round-check. [line 76]" ' vla:3812
            Exit Function
        Case 44
            vla_step_text = "Set thousand-check to 1,000 plus 500. [line 77]" ' vla:3813
            Exit Function
        Case 45
            vla_step_text = "Log ""thousands read as "" joined with thousand-check. [line 78]" ' vla:3814
            Exit Function
        Case 46
            vla_step_text = "Log ""probe is "" joined with probe. [line 79]" ' vla:3815
            Exit Function
        Case 47
            vla_step_text = "Create a number called countdown. [line 81]" ' vla:3816
            Exit Function
        Case 48
            vla_step_text = "Set countdown to 3. [line 82]" ' vla:3817
            Exit Function
        Case 49
            vla_step_text = "While countdown is greater than 0: [line 83]" ' vla:3818
            Exit Function
        Case 50
            vla_step_text = "Log ""countdown "" joined with countdown. [line 84]" ' vla:3819
            Exit Function
        Case 51
            vla_step_text = "Decrease countdown by 1. [line 85]" ' vla:3820
            Exit Function
        Case 52
            vla_step_text = "Put value into column F row row-number. [line 90]" ' vla:3821
            Exit Function
        Case 53
            vla_step_text = "Stamp. [line 92]" ' vla:3822
            Exit Function
        Case 54
            vla_step_text = "Stamp with row-number of 2 and value of ""beta"". [line 93]" ' vla:3823
            Exit Function
        Case 55
            vla_step_text = "Set f-last to last filled row of column F. [line 94]" ' vla:3824
            Exit Function
        Case 56
            vla_step_text = "Log ""column F filled to row "" joined with f-last. [line 95]" ' vla:3825
            Exit Function
        Case 57
            vla_step_text = "Count echo-row from 1 to f-last, log ""echo "" joined with echo-row. [line 96]" ' vla:3826
            Exit Function
        Case 58
            vla_step_text = "Count check-row from 1 to f-last: [line 97]" ' vla:3827
            Exit Function
        Case 59
            vla_step_text = "If cell in column F row check-row contains ""ok"", make cell in column F row check-row bold. [line 98]" ' vla:3828
            Exit Function
        Case 60
            vla_step_text = "Count stripe-row from 1 to f-last step 2: [line 103]" ' vla:3829
            Exit Function
        Case 61
            vla_step_text = "Make cell in column F row stripe-row bold. [line 104]" ' vla:3830
            Exit Function
        Case 62
            vla_step_text = "Count back-row down from f-last to 1 step 2: [line 106]" ' vla:3831
            Exit Function
        Case 63
            vla_step_text = "Log ""back-row "" joined with back-row. [line 107]" ' vla:3832
            Exit Function
        Case 64
            vla_step_text = "Count search-row from 1 to 10: [line 111]" ' vla:3833
            Exit Function
        Case 65
            vla_step_text = "If search-row is 3, stop the loop. [line 112]" ' vla:3834
            Exit Function
        Case 66
            vla_step_text = "Log ""stopped at "" joined with search-row. [line 114]" ' vla:3835
            Exit Function
        Case 67
            vla_step_text = "Create a list called found-items. [line 119]" ' vla:3836
            Exit Function
        Case 68
            vla_step_text = "Repeat 4 times: [line 120]" ' vla:3837
            Exit Function
        Case 69
            vla_step_text = "If counter is greater than 2, append counter times 100 to found-items. [line 121]" ' vla:3838
            Exit Function
        Case 70
            vla_step_text = "For each f in found-items, log ""found "" joined with f. [line 123]" ' vla:3839
            Exit Function
        Case 71
            vla_step_text = "Set list-count to count of found-items. [line 124]" ' vla:3840
            Exit Function
        Case 72
            vla_step_text = "Log ""list holds "" joined with list-count. [line 125]" ' vla:3841
            Exit Function
        Case 73
            vla_step_text = "Create a text called verdict. [line 131]" ' vla:3842
            Exit Function
        Case 74
            vla_step_text = "If grand is greater than 100: [line 132]" ' vla:3843
            Exit Function
        Case 75
            vla_step_text = "Set verdict to ""huge"". [line 133]" ' vla:3844
            Exit Function
        Case 76
            vla_step_text = "Set verdict to ""solid"". [line 136]" ' vla:3845
            Exit Function
        Case 77
            vla_step_text = "Set verdict to ""small"". [line 139]" ' vla:3846
            Exit Function
        Case 78
            vla_step_text = "Create a text called region-label. [line 141]" ' vla:3847
            Exit Function
        Case 79
            vla_step_text = "Set region to ""South"". [line 142]" ' vla:3848
            Exit Function
        Case 80
            vla_step_text = "When region is ""North"": [line 143]" ' vla:3849
            Exit Function
        Case 81
            vla_step_text = "Set region-label to ""cold"". [line 144]" ' vla:3850
            Exit Function
        Case 82
            vla_step_text = "Set region-label to ""warm"". [line 147]" ' vla:3851
            Exit Function
        Case 83
            vla_step_text = "Set region-label to ""unknown"". [line 150]" ' vla:3852
            Exit Function
        Case 84
            vla_step_text = "Log ""verdict "" joined with verdict joined with "", region "" joined with region-label. [line 152]" ' vla:3853
            Exit Function
        Case 85
            vla_step_text = "Create a number called until-count. [line 154]" ' vla:3854
            Exit Function
        Case 86
            vla_step_text = "Set fuel to 3. [line 155]" ' vla:3855
            Exit Function
        Case 87
            vla_step_text = "Repeat until fuel is 0: [line 156]" ' vla:3856
            Exit Function
        Case 88
            vla_step_text = "Decrease fuel by 1. [line 157]" ' vla:3857
            Exit Function
        Case 89
            vla_step_text = "Increase until-count by 1. [line 158]" ' vla:3858
            Exit Function
        Case 90
            vla_step_text = "Log ""repeat-until ran "" joined with until-count joined with "" times"". [line 160]" ' vla:3859
            Exit Function
        Case 91
            vla_step_text = "Create a text called rescue. [line 166]" ' vla:3860
            Exit Function
        Case 92
            vla_step_text = "Try: [line 167]" ' vla:3861
            Exit Function
        Case 93
            vla_step_text = "Go to sheet Nowhere-Land. [line 168]" ' vla:3862
            Exit Function
        Case 94
            vla_step_text = "Set rescue to ""unreachable"". [line 169]" ' vla:3863
            Exit Function
        Case 95
            vla_step_text = "Log ""the problem was "" joined with the problem. [line 172]" ' vla:3864
            Exit Function
        Case 96
            vla_step_text = "If the problem is not empty, set rescue to ""rescued"". [line 173]" ' vla:3865
            Exit Function
        Case 97
            vla_step_text = "Create a number called risk-free. [line 175]" ' vla:3866
            Exit Function
        Case 98
            vla_step_text = "Try: [line 176]" ' vla:3867
            Exit Function
        Case 99
            vla_step_text = "Set risk-free to 7. [line 177]" ' vla:3868
            Exit Function
        Case 100
            vla_step_text = "Set risk-free to -1. [line 180]" ' vla:3869
            Exit Function
        Case 101
            vla_step_text = "Log ""rescue "" joined with rescue joined with "", risk-free "" joined with risk-free. [line 182]" ' vla:3870
            Exit Function
        Case 102
            vla_step_text = "Give back amount times 0.08. [line 187]" ' vla:3871
            Exit Function
        Case 103
            vla_step_text = "Set fee to tax of 100. [line 189]" ' vla:3872
            Exit Function
        Case 104
            vla_step_text = "Log ""fee is "" joined with fee. [line 190]" ' vla:3873
            Exit Function
        Case 105
            vla_step_text = "Create a text called fee-size. [line 191]" ' vla:3874
            Exit Function
        Case 106
            vla_step_text = "If tax of 50 is greater than 3: [line 192]" ' vla:3875
            Exit Function
        Case 107
            vla_step_text = "Set fee-size to ""big"". [line 193]" ' vla:3876
            Exit Function
        Case 108
            vla_step_text = "Set fee-size to ""small"". [line 196]" ' vla:3877
            Exit Function
        Case 109
            vla_step_text = "Fit column A. [line 199]" ' vla:3878
            Exit Function
        Case 110
            vla_step_text = "Fit column B. [line 200]" ' vla:3879
            Exit Function
        Case 111
            vla_step_text = "Set font-color of cell A1 to hot-pink. [line 201]" ' vla:3880
            Exit Function
        Case 112
            vla_step_text = "Give back sale times rate. [line 206]" ' vla:3881
            Exit Function
        Case 113
            vla_step_text = "Set full-commission to commission using sale of 2000 and rate of 10%. [line 208]" ' vla:3882
            Exit Function
        Case 114
            vla_step_text = "Set default-commission to get commission using sale of 600. [line 209]" ' vla:3883
            Exit Function
        Case 115
            vla_step_text = "Log ""commissions "" joined with full-commission joined with "" / "" joined with default-commission. [line 210]" ' vla:3884
            Exit Function
        Case 116
            vla_step_text = "Try: [line 219]" ' vla:3885
            Exit Function
        Case 117
            vla_step_text = "Add sheet called ""Q1 Data"". [line 220]" ' vla:3886
            Exit Function
        Case 118
            vla_step_text = "Go to sheet Output. [line 222]" ' vla:3887
            Exit Function
        Case 119
            vla_step_text = "Put ""spaced"" into cell 'Q1 Data'!A1. [line 223]" ' vla:3888
            Exit Function
        Case 120
            vla_step_text = "Create a text called spaced-check. [line 224]" ' vla:3889
            Exit Function
        Case 121
            vla_step_text = "Set spaced-check to value in cell 'Q1 Data'!A1. [line 225]" ' vla:3890
            Exit Function
        Case 122
            vla_step_text = "Give back 20%. [line 230]" ' vla:3891
            Exit Function
        Case 123
            vla_step_text = "Create a number called growth-check. [line 237]" ' vla:3892
            Exit Function
        Case 124
            vla_step_text = "Set growth-check to 200. [line 238]" ' vla:3893
            Exit Function
        Case 125
            vla_step_text = "Grow growth-check by 10%. [line 239]" ' vla:3894
            Exit Function
        Case 126
            vla_step_text = "Increase growth-check by 50%. [line 240]" ' vla:3895
            Exit Function
        Case 127
            vla_step_text = "Add 100 percent to growth-check. [line 241]" ' vla:3896
            Exit Function
        Case 128
            vla_step_text = "Decrease growth-check by 75%. [line 242]" ' vla:3897
            Exit Function
        Case 129
            vla_step_text = "Set pick-check to item 2 of found-items. [line 243]" ' vla:3898
            Exit Function
        Case 130
            vla_step_text = "Log ""grew to "" joined with growth-check joined with "", picked "" joined with pick-check joined with "", first "" joined with first of found-items. [line 244]" ' vla:3899
            Exit Function
        Case 131
            vla_step_text = "Create a text called quote-check. [line 245]" ' vla:3900
            Exit Function
        Case 132
            vla_step_text = "Set quote-check to ""He said """"ok"""""". [line 246]" ' vla:3901
            Exit Function
        Case 133
            vla_step_text = "Create a lookup called prices. [line 260]" ' vla:3902
            Exit Function
        Case 134
            vla_step_text = "Store 100 at key ""ax-7"" in prices. [line 261]" ' vla:3903
            Exit Function
        Case 135
            vla_step_text = "Store 250 under ""bx-2"" in prices. [line 262]" ' vla:3904
            Exit Function
        Case 136
            vla_step_text = "Store 120 at key ""AX-7"" in prices. [line 263]" ' vla:3905
            Exit Function
        Case 137
            vla_step_text = "Set ax-price to prices for ""ax-7"". [line 264]" ' vla:3906
            Exit Function
        Case 138
            vla_step_text = "Set key-count to count of keys of prices. [line 265]" ' vla:3907
            Exit Function
        Case 139
            vla_step_text = "Set price-sum to prices for ""ax-7"" plus prices for ""bx-2"". [line 266]" ' vla:3908
            Exit Function
        Case 140
            vla_step_text = "Create a text called key-list. [line 267]" ' vla:3909
            Exit Function
        Case 141
            vla_step_text = "For each k in keys of prices: [line 268]" ' vla:3910
            Exit Function
        Case 142
            vla_step_text = "Set key-list to key-list joined with k. [line 269]" ' vla:3911
            Exit Function
        Case 143
            vla_step_text = "Create a text called price-verdict. [line 271]" ' vla:3912
            Exit Function
        Case 144
            vla_step_text = "If prices for ""bx-2"" is greater than 200, set price-verdict to ""steep"". [line 272]" ' vla:3913
            Exit Function
        Case 145
            vla_step_text = "Create a text called pair-trace. [line 273]" ' vla:3914
            Exit Function
        Case 146
            vla_step_text = "For each pair in prices: [line 274]" ' vla:3915
            Exit Function
        Case 147
            vla_step_text = "If value of pair is greater than 200, set pair-trace to pair-trace joined with key of pair. [line 275]" ' vla:3916
            Exit Function
        Case 148
            vla_step_text = "Log ""lookup: ax "" joined with ax-price joined with "", keys "" joined with key-list joined with "", pairs "" joined with pair-trace. [line 277]" ' vla:3917
            Exit Function
        Case 149
            vla_step_text = "Put ""Item"" into cell J1. [line 284]" ' vla:3918
            Exit Function
        Case 150
            vla_step_text = "Put ""Amount"" into cell K1. [line 285]" ' vla:3919
            Exit Function
        Case 151
            vla_step_text = "Put ""Widget"" into cell J2. [line 286]" ' vla:3920
            Exit Function
        Case 152
            vla_step_text = "Put 10 into cell K2. [line 287]" ' vla:3921
            Exit Function
        Case 153
            vla_step_text = "Put ""Gadget"" into cell J3. [line 288]" ' vla:3922
            Exit Function
        Case 154
            vla_step_text = "Put 20 into cell K3. [line 289]" ' vla:3923
            Exit Function
        Case 155
            vla_step_text = "Turn J1:K3 into a table called SalesTable. [line 290]" ' vla:3924
            Exit Function
        Case 156
            vla_step_text = "Set style of table SalesTable to TableStyleMedium9. [line 291]" ' vla:3925
            Exit Function
        Case 157
            vla_step_text = "Show the total row of table SalesTable. [line 292]" ' vla:3926
            Exit Function
        Case 158
            vla_step_text = "Put ""X"" into cell J5. [line 296]" ' vla:3927
            Exit Function
        Case 159
            vla_step_text = "Put ""Y"" into cell J6. [line 297]" ' vla:3928
            Exit Function
        Case 160
            vla_step_text = "Turn J5:J6 into a table called QuietTable. [line 298]" ' vla:3929
            Exit Function
        Case 161
            vla_step_text = "Show the total row of table QuietTable. [line 299]" ' vla:3930
            Exit Function
        Case 162
            vla_step_text = "Hide the total row of table QuietTable. [line 300]" ' vla:3931
            Exit Function
        Case 163
            vla_step_text = "Put ""A"" into cell J8. [line 304]" ' vla:3932
            Exit Function
        Case 164
            vla_step_text = "Put ""B"" into cell J9. [line 305]" ' vla:3933
            Exit Function
        Case 165
            vla_step_text = "Turn J8:J9 into a table called TempTable. [line 306]" ' vla:3934
            Exit Function
        Case 166
            vla_step_text = "Turn table TempTable back into a range. [line 307]" ' vla:3935
            Exit Function
        Case 167
            vla_step_text = "Put ""Item"" into cell J11. [line 312]" ' vla:3936
            Exit Function
        Case 168
            vla_step_text = "Put ""Qty"" into cell K11. [line 313]" ' vla:3937
            Exit Function
        Case 169
            vla_step_text = "Put ""Bolt"" into cell J12. [line 314]" ' vla:3938
            Exit Function
        Case 170
            vla_step_text = "Put 5 into cell K12. [line 315]" ' vla:3939
            Exit Function
        Case 171
            vla_step_text = "Put ""Nut"" into cell J13. [line 316]" ' vla:3940
            Exit Function
        Case 172
            vla_step_text = "Put 8 into cell K13. [line 317]" ' vla:3941
            Exit Function
        Case 173
            vla_step_text = "Turn J11:K13 into a table called EditTable. [line 318]" ' vla:3942
            Exit Function
        Case 174
            vla_step_text = "Add a row to table EditTable. [line 319]" ' vla:3943
            Exit Function
        Case 175
            vla_step_text = "Delete row 1 of table EditTable. [line 320]" ' vla:3944
            Exit Function
        Case 176
            vla_step_text = "Set qty-values to column Qty of table EditTable. [line 321]" ' vla:3945
            Exit Function
        Case 177
            vla_step_text = "For each q in qty-values, log q. [line 322]" ' vla:3946
            Exit Function
        Case 178
            vla_step_text = "Put ""Region"" into cell N1. [line 332]" ' vla:3947
            Exit Function
        Case 179
            vla_step_text = "Put ""Product"" into cell O1. [line 333]" ' vla:3948
            Exit Function
        Case 180
            vla_step_text = "Put ""Segment"" into cell P1. [line 334]" ' vla:3949
            Exit Function
        Case 181
            vla_step_text = "Put ""Channel"" into cell Q1. [line 335]" ' vla:3950
            Exit Function
        Case 182
            vla_step_text = "Put ""Units"" into cell R1. [line 336]" ' vla:3951
            Exit Function
        Case 183
            vla_step_text = "Put ""Revenue"" into cell S1. [line 337]" ' vla:3952
            Exit Function
        Case 184
            vla_step_text = "Put ""North"" into cell N2. [line 338]" ' vla:3953
            Exit Function
        Case 185
            vla_step_text = "Put ""Widget"" into cell O2. [line 339]" ' vla:3954
            Exit Function
        Case 186
            vla_step_text = "Put ""Retail"" into cell P2. [line 340]" ' vla:3955
            Exit Function
        Case 187
            vla_step_text = "Put ""Online"" into cell Q2. [line 341]" ' vla:3956
            Exit Function
        Case 188
            vla_step_text = "Put 10 into cell R2. [line 342]" ' vla:3957
            Exit Function
        Case 189
            vla_step_text = "Put 500 into cell S2. [line 343]" ' vla:3958
            Exit Function
        Case 190
            vla_step_text = "Put ""North"" into cell N3. [line 344]" ' vla:3959
            Exit Function
        Case 191
            vla_step_text = "Put ""Gadget"" into cell O3. [line 345]" ' vla:3960
            Exit Function
        Case 192
            vla_step_text = "Put ""Wholesale"" into cell P3. [line 346]" ' vla:3961
            Exit Function
        Case 193
            vla_step_text = "Put ""Store"" into cell Q3. [line 347]" ' vla:3962
            Exit Function
        Case 194
            vla_step_text = "Put 5 into cell R3. [line 348]" ' vla:3963
            Exit Function
        Case 195
            vla_step_text = "Put 200 into cell S3. [line 349]" ' vla:3964
            Exit Function
        Case 196
            vla_step_text = "Put ""South"" into cell N4. [line 350]" ' vla:3965
            Exit Function
        Case 197
            vla_step_text = "Put ""Widget"" into cell O4. [line 351]" ' vla:3966
            Exit Function
        Case 198
            vla_step_text = "Put ""Wholesale"" into cell P4. [line 352]" ' vla:3967
            Exit Function
        Case 199
            vla_step_text = "Put ""Online"" into cell Q4. [line 353]" ' vla:3968
            Exit Function
        Case 200
            vla_step_text = "Put 20 into cell R4. [line 354]" ' vla:3969
            Exit Function
        Case 201
            vla_step_text = "Put 900 into cell S4. [line 355]" ' vla:3970
            Exit Function
        Case 202
            vla_step_text = "Put ""South"" into cell N5. [line 356]" ' vla:3971
            Exit Function
        Case 203
            vla_step_text = "Put ""Gadget"" into cell O5. [line 357]" ' vla:3972
            Exit Function
        Case 204
            vla_step_text = "Put ""Retail"" into cell P5. [line 358]" ' vla:3973
            Exit Function
        Case 205
            vla_step_text = "Put ""Store"" into cell Q5. [line 359]" ' vla:3974
            Exit Function
        Case 206
            vla_step_text = "Put 8 into cell R5. [line 360]" ' vla:3975
            Exit Function
        Case 207
            vla_step_text = "Put 300 into cell S5. [line 361]" ' vla:3976
            Exit Function
        Case 208
            vla_step_text = "Put ""East"" into cell N6. [line 362]" ' vla:3977
            Exit Function
        Case 209
            vla_step_text = "Put ""Widget"" into cell O6. [line 363]" ' vla:3978
            Exit Function
        Case 210
            vla_step_text = "Put ""Retail"" into cell P6. [line 364]" ' vla:3979
            Exit Function
        Case 211
            vla_step_text = "Put ""Online"" into cell Q6. [line 365]" ' vla:3980
            Exit Function
        Case 212
            vla_step_text = "Put 12 into cell R6. [line 366]" ' vla:3981
            Exit Function
        Case 213
            vla_step_text = "Put 600 into cell S6. [line 367]" ' vla:3982
            Exit Function
        Case 214
            vla_step_text = "Put ""East"" into cell N7. [line 368]" ' vla:3983
            Exit Function
        Case 215
            vla_step_text = "Put ""Gadget"" into cell O7. [line 369]" ' vla:3984
            Exit Function
        Case 216
            vla_step_text = "Put ""Wholesale"" into cell P7. [line 370]" ' vla:3985
            Exit Function
        Case 217
            vla_step_text = "Put ""Store"" into cell Q7. [line 371]" ' vla:3986
            Exit Function
        Case 218
            vla_step_text = "Put 6 into cell R7. [line 372]" ' vla:3987
            Exit Function
        Case 219
            vla_step_text = "Put 250 into cell S7. [line 373]" ' vla:3988
            Exit Function
        Case 220
            vla_step_text = "Make a pivot table from N1:S7 at U1 called SalesPivot. [line 381]" ' vla:3989
            Exit Function
        Case 221
            vla_step_text = "Add rows of Region, Product to pivot SalesPivot. [line 382]" ' vla:3990
            Exit Function
        Case 222
            vla_step_text = "Add columns of Segment to pivot SalesPivot. [line 383]" ' vla:3991
            Exit Function
        Case 223
            vla_step_text = "Add filters of Channel to pivot SalesPivot. [line 384]" ' vla:3992
            Exit Function
        Case 224
            vla_step_text = "Add Revenue to pivot SalesPivot as a sum. [line 392]" ' vla:3993
            Exit Function
        Case 225
            vla_step_text = "Add Revenue, Units to pivot SalesPivot as a count. [line 393]" ' vla:3994
            Exit Function
        Case 226
            vla_step_text = "Add Units to pivot SalesPivot as an average. [line 394]" ' vla:3995
            Exit Function
        Case 227
            vla_step_text = "Make a pivot table from N1:S7 at N20 called FullPivot with rows of Region, Product, and Channel and columns of Segment and values of Revenue, Units. [line 406]" ' vla:3996
            Exit Function
        Case 228
            vla_step_text = "Refresh pivot SalesPivot. [line 415]" ' vla:3997
            Exit Function
        Case 229
            vla_step_text = "Refresh every pivot table. [line 416]" ' vla:3998
            Exit Function
        Case 230
            vla_step_text = "Collapse Region in pivot SalesPivot. [line 417]" ' vla:3999
            Exit Function
        Case 231
            vla_step_text = "Collapse Product in pivot FullPivot. [line 427]" ' vla:4000
            Exit Function
        Case 232
            vla_step_text = "Expand Product in pivot FullPivot. [line 428]" ' vla:4001
            Exit Function
        Case 233
            vla_step_text = "Show pivot SalesPivot in tabular form. [line 443]" ' vla:4002
            Exit Function
        Case 234
            vla_step_text = "Show pivot SalesPivot in compact form. [line 444]" ' vla:4003
            Exit Function
        Case 235
            vla_step_text = "Show pivot FullPivot in outline form. [line 445]" ' vla:4004
            Exit Function
        Case 236
            vla_step_text = "Hide subtotals for Region in pivot SalesPivot. [line 453]" ' vla:4005
            Exit Function
        Case 237
            vla_step_text = "Hide subtotals for Product in pivot FullPivot. [line 454]" ' vla:4006
            Exit Function
        Case 238
            vla_step_text = "Show subtotals for Product in pivot FullPivot. [line 455]" ' vla:4007
            Exit Function
        Case 239
            vla_step_text = "Add a blank row after Region in pivot SalesPivot. [line 457]" ' vla:4008
            Exit Function
        Case 240
            vla_step_text = "Add a blank row after Product in pivot FullPivot. [line 458]" ' vla:4009
            Exit Function
        Case 241
            vla_step_text = "Remove a blank row after Product in pivot FullPivot. [line 459]" ' vla:4010
            Exit Function
        Case 242
            vla_step_text = "Sort Region in pivot SalesPivot descending. [line 476]" ' vla:4011
            Exit Function
        Case 243
            vla_step_text = "Sort Region in pivot SalesPivot ascending. [line 477]" ' vla:4012
            Exit Function
        Case 244
            vla_step_text = "Sort Product in pivot FullPivot descending by Revenue. [line 478]" ' vla:4013
            Exit Function
        Case 245
            vla_step_text = "Sort Product in pivot FullPivot ascending by Revenue. [line 479]" ' vla:4014
            Exit Function
        Case 246
            vla_step_text = "Make a pivot table from N1:S7 at N200 called RenameMePivot. [line 490]" ' vla:4015
            Exit Function
        Case 247
            vla_step_text = "Rename pivot RenameMePivot to RenamedPivot. [line 491]" ' vla:4016
            Exit Function
        Case 248
            vla_step_text = "Make a pivot table from N1:S7 at N220 called ClearMePivot. [line 493]" ' vla:4017
            Exit Function
        Case 249
            vla_step_text = "Add rows of Region to pivot ClearMePivot. [line 494]" ' vla:4018
            Exit Function
        Case 250
            vla_step_text = "Add Revenue to pivot ClearMePivot as a sum. [line 495]" ' vla:4019
            Exit Function
        Case 251
            vla_step_text = "Clear pivot ClearMePivot. [line 496]" ' vla:4020
            Exit Function
        Case 252
            vla_step_text = "Make a pivot table from N1:S7 at N240 called RemoveFieldMePivot. [line 498]" ' vla:4021
            Exit Function
        Case 253
            vla_step_text = "Add rows of Region, Product to pivot RemoveFieldMePivot. [line 499]" ' vla:4022
            Exit Function
        Case 254
            vla_step_text = "Remove Product from pivot RemoveFieldMePivot. [line 500]" ' vla:4023
            Exit Function
        Case 255
            vla_step_text = "Put ""West"" into cell N8. [line 506]" ' vla:4024
            Exit Function
        Case 256
            vla_step_text = "Put ""Widget"" into cell O8. [line 507]" ' vla:4025
            Exit Function
        Case 257
            vla_step_text = "Put ""Retail"" into cell P8. [line 508]" ' vla:4026
            Exit Function
        Case 258
            vla_step_text = "Put ""Online"" into cell Q8. [line 509]" ' vla:4027
            Exit Function
        Case 259
            vla_step_text = "Put 15 into cell R8. [line 510]" ' vla:4028
            Exit Function
        Case 260
            vla_step_text = "Put 700 into cell S8. [line 511]" ' vla:4029
            Exit Function
        Case 261
            vla_step_text = "Make a pivot table from N1:S7 at N260 called SourceTestPivot. [line 513]" ' vla:4030
            Exit Function
        Case 262
            vla_step_text = "Add rows of Region to pivot SourceTestPivot. [line 514]" ' vla:4031
            Exit Function
        Case 263
            vla_step_text = "Change the source of pivot SourceTestPivot to N1:S8. [line 515]" ' vla:4032
            Exit Function
        Case 264
            vla_step_text = "Make a pivot table from N1:S7 at N40 called TempPivot. [line 520]" ' vla:4033
            Exit Function
        Case 265
            vla_step_text = "Delete pivot TempPivot. [line 521]" ' vla:4034
            Exit Function
        Case 266
            vla_step_text = "Put grand into cell H1. [line 524]" ' vla:4035
            Exit Function
        Case 267
            vla_step_text = "Put biggest into cell H2. [line 525]" ' vla:4036
            Exit Function
        Case 268
            vla_step_text = "Put round-check into cell H3. [line 526]" ' vla:4037
            Exit Function
        Case 269
            vla_step_text = "Put thousand-check into cell H4. [line 527]" ' vla:4038
            Exit Function
        Case 270
            vla_step_text = "Put search-row into cell H5. [line 528]" ' vla:4039
            Exit Function
        Case 271
            vla_step_text = "Put f-last into cell H6. [line 529]" ' vla:4040
            Exit Function
        Case 272
            vla_step_text = "Put list-count into cell H7. [line 530]" ' vla:4041
            Exit Function
        Case 273
            vla_step_text = "Put verdict into cell H8. [line 531]" ' vla:4042
            Exit Function
        Case 274
            vla_step_text = "Put region-label into cell H9. [line 532]" ' vla:4043
            Exit Function
        Case 275
            vla_step_text = "Put until-count into cell H10. [line 533]" ' vla:4044
            Exit Function
        Case 276
            vla_step_text = "Put rescue into cell H11. [line 534]" ' vla:4045
            Exit Function
        Case 277
            vla_step_text = "Put risk-free into cell H12. [line 535]" ' vla:4046
            Exit Function
        Case 278
            vla_step_text = "Put fee into cell H13. [line 536]" ' vla:4047
            Exit Function
        Case 279
            vla_step_text = "Put fee-size into cell H14. [line 537]" ' vla:4048
            Exit Function
        Case 280
            vla_step_text = "Put num-col-check into cell H15. [line 538]" ' vla:4049
            Exit Function
        Case 281
            vla_step_text = "Put value-word-check into cell H17. [line 539]" ' vla:4050
            Exit Function
        Case 282
            vla_step_text = "Put spaced-check into cell H18. [line 540]" ' vla:4051
            Exit Function
        Case 283
            vla_step_text = "Put full-commission into cell H19. [line 541]" ' vla:4052
            Exit Function
        Case 284
            vla_step_text = "Put default-commission into cell H20. [line 542]" ' vla:4053
            Exit Function
        Case 285
            vla_step_text = "Put growth-check into cell H21. [line 543]" ' vla:4054
            Exit Function
        Case 286
            vla_step_text = "Put pick-check into cell H22. [line 544]" ' vla:4055
            Exit Function
        Case 287
            vla_step_text = "Put quote-check into cell H23. [line 545]" ' vla:4056
            Exit Function
        Case 288
            vla_step_text = "Put vat-rate into cell H24. [line 546]" ' vla:4057
            Exit Function
        Case 289
            vla_step_text = "Put ax-price into cell H26. [line 547]" ' vla:4058
            Exit Function
        Case 290
            vla_step_text = "Put key-count into cell H27. [line 548]" ' vla:4059
            Exit Function
        Case 291
            vla_step_text = "Put key-list into cell H28. [line 549]" ' vla:4060
            Exit Function
        Case 292
            vla_step_text = "Put price-sum into cell H29. [line 550]" ' vla:4061
            Exit Function
        Case 293
            vla_step_text = "Put price-verdict into cell H30. [line 551]" ' vla:4062
            Exit Function
        Case 294
            vla_step_text = "Put pair-trace into cell H31. [line 552]" ' vla:4063
            Exit Function
        Case 295
            vla_step_text = "(set! (range ""h25"") ""vla-row"") [line 559]" ' vla:4064
            Exit Function
        Case 296
            vla_step_text = "Repeat 2 times: [line 560]" ' vla:4065
            Exit Function
        Case 297
            vla_step_text = "(debug-print counter) [line 561]" ' vla:4066
            Exit Function
        Case 298
            vla_step_text = "Work on sheet Demo. [line 576]" ' vla:4067
            Exit Function
        Case 299
            vla_step_text = "Make cell A1 italic. [line 579]" ' vla:4068
            Exit Function
        Case 300
            vla_step_text = "Make cell A2 red. [line 580]" ' vla:4069
            Exit Function
        Case 301
            vla_step_text = "Make cell A3 yellow. [line 581]" ' vla:4070
            Exit Function
        Case 302
            vla_step_text = "Set font-color of cell A4 to hot-pink. [line 582]" ' vla:4071
            Exit Function
        Case 303
            vla_step_text = "Set fill-color of cell A5 to ""#FF69B4"". [line 583]" ' vla:4072
            Exit Function
        Case 304
            vla_step_text = "Clear color of cell A5. [line 584]" ' vla:4073
            Exit Function
        Case 305
            vla_step_text = "Add border to range A1:E10. [line 585]" ' vla:4074
            Exit Function
        Case 306
            vla_step_text = "Format cell B1 as currency. [line 586]" ' vla:4075
            Exit Function
        Case 307
            vla_step_text = "Format cell B2 as percent. [line 587]" ' vla:4076
            Exit Function
        Case 308
            vla_step_text = "Format cell B3 as date. [line 588]" ' vla:4077
            Exit Function
        Case 309
            vla_step_text = "Center cell A1. [line 589]" ' vla:4078
            Exit Function
        Case 310
            vla_step_text = "Center range A1:E1. [line 590]" ' vla:4079
            Exit Function
        Case 311
            vla_step_text = "Align range A2:A5 left. [line 591]" ' vla:4080
            Exit Function
        Case 312
            vla_step_text = "Align cell A6 center. [line 592]" ' vla:4081
            Exit Function
        Case 313
            vla_step_text = "Wrap text in range C1:C5. [line 593]" ' vla:4082
            Exit Function
        Case 314
            vla_step_text = "Unwrap text in range C1:C5. [line 594]" ' vla:4083
            Exit Function
        Case 315
            vla_step_text = "Merge range D1:D3. [line 595]" ' vla:4084
            Exit Function
        Case 316
            vla_step_text = "Unmerge range D1:D3. [line 596]" ' vla:4085
            Exit Function
        Case 317
            vla_step_text = "Set height of row 1 to 30. [line 601]" ' vla:4086
            Exit Function
        Case 318
            vla_step_text = "Set width of column A to 20. [line 602]" ' vla:4087
            Exit Function
        Case 319
            vla_step_text = "Insert column before B. [line 603]" ' vla:4088
            Exit Function
        Case 320
            vla_step_text = "Hide column C. [line 604]" ' vla:4089
            Exit Function
        Case 321
            vla_step_text = "Unhide column C. [line 605]" ' vla:4090
            Exit Function
        Case 322
            vla_step_text = "Delete column B. [line 606]" ' vla:4091
            Exit Function
        Case 323
            vla_step_text = "Insert row at 5. [line 607]" ' vla:4092
            Exit Function
        Case 324
            vla_step_text = "Insert a row. [line 608]" ' vla:4093
            Exit Function
        Case 325
            vla_step_text = "Delete row 2. [line 609]" ' vla:4094
            Exit Function
        Case 326
            vla_step_text = "Delete the third row. [line 610]" ' vla:4095
            Exit Function
        Case 327
            vla_step_text = "Hide row 10. [line 611]" ' vla:4096
            Exit Function
        Case 328
            vla_step_text = "Unhide row 10. [line 612]" ' vla:4097
            Exit Function
        Case 329
            vla_step_text = "Freeze top row. [line 613]" ' vla:4098
            Exit Function
        Case 330
            vla_step_text = "Unfreeze panes. [line 614]" ' vla:4099
            Exit Function
        Case 331
            vla_step_text = "Fit column A. [line 615]" ' vla:4100
            Exit Function
        Case 332
            vla_step_text = "Fit all columns. [line 616]" ' vla:4101
            Exit Function
        Case 333
            vla_step_text = "Put ""Region"" into cell G1. [line 620]" ' vla:4102
            Exit Function
        Case 334
            vla_step_text = "Put ""Amount"" into cell H1. [line 621]" ' vla:4103
            Exit Function
        Case 335
            vla_step_text = "Put ""Notes"" into cell I1. [line 622]" ' vla:4104
            Exit Function
        Case 336
            vla_step_text = "Put ""West"" into cell G2. [line 623]" ' vla:4105
            Exit Function
        Case 337
            vla_step_text = "Put 100 into cell H2. [line 624]" ' vla:4106
            Exit Function
        Case 338
            vla_step_text = "Put ""ok"" into cell I2. [line 625]" ' vla:4107
            Exit Function
        Case 339
            vla_step_text = "Put ""East"" into cell G3. [line 626]" ' vla:4108
            Exit Function
        Case 340
            vla_step_text = "Put 250 into cell H3. [line 627]" ' vla:4109
            Exit Function
        Case 341
            vla_step_text = "Put ""ok"" into cell I3. [line 628]" ' vla:4110
            Exit Function
        Case 342
            vla_step_text = "Put ""West"" into cell G4. [line 629]" ' vla:4111
            Exit Function
        Case 343
            vla_step_text = "Put 100 into cell H4. [line 630]" ' vla:4112
            Exit Function
        Case 344
            vla_step_text = "Put ""dup"" into cell I4. [line 631]" ' vla:4113
            Exit Function
        Case 345
            vla_step_text = "Sort range G1:I4 by column H1 descending. [line 632]" ' vla:4114
            Exit Function
        Case 346
            vla_step_text = "Keep only rows of range G1:I4 where column 1 is ""West"". [line 633]" ' vla:4115
            Exit Function
        Case 347
            vla_step_text = "Show all rows. [line 634]" ' vla:4116
            Exit Function
        Case 348
            vla_step_text = "Replace ""dup"" with ""ok"" in range G1:I4. [line 635]" ' vla:4117
            Exit Function
        Case 349
            vla_step_text = "Remove duplicates from range G1:I4 by column 1. [line 636]" ' vla:4118
            Exit Function
        Case 350
            vla_step_text = "Replace ""ok"" with ""fine"" in column I. [line 637]" ' vla:4119
            Exit Function
        Case 351
            vla_step_text = "Convert range G1:I4 to values. [line 638]" ' vla:4120
            Exit Function
        Case 352
            vla_step_text = "Clear formatting of range G1:I4. [line 639]" ' vla:4121
            Exit Function
        Case 353
            vla_step_text = "Name range G1:I4 as demo_table. [line 640]" ' vla:4122
            Exit Function
        Case 354
            vla_step_text = "Copy range G1:I4 to range K1:M4. [line 641]" ' vla:4123
            Exit Function
        Case 355
            vla_step_text = "Set tab-color of sheet Demo to hot-pink. [line 645]" ' vla:4124
            Exit Function
        Case 356
            vla_step_text = "Protect this sheet with password ""demo123"". [line 646]" ' vla:4125
            Exit Function
        Case 357
            vla_step_text = "Unprotect this sheet with password ""demo123"". [line 647]" ' vla:4126
            Exit Function
        Case 358
            vla_step_text = "Try: [line 675]" ' vla:4127
            Exit Function
        Case 359
            vla_step_text = "Delete sheet GStruct. [line 676]" ' vla:4128
            Exit Function
        Case 360
            vla_step_text = "Work on sheet GStruct. [line 678]" ' vla:4129
            Exit Function
        Case 361
            vla_step_text = "Hide row 3. [line 680]" ' vla:4130
            Exit Function
        Case 362
            vla_step_text = "Hide column B. [line 681]" ' vla:4131
            Exit Function
        Case 363
            vla_step_text = "Unhide all rows and columns. [line 682]" ' vla:4132
            Exit Function
        Case 364
            vla_step_text = "Set font size of cell A6 to 36. [line 684]" ' vla:4133
            Exit Function
        Case 365
            vla_step_text = "Fit row 6. [line 685]" ' vla:4134
            Exit Function
        Case 366
            vla_step_text = "Group rows 10 through 12. [line 687]" ' vla:4135
            Exit Function
        Case 367
            vla_step_text = "Group rows 14 through 16. [line 688]" ' vla:4136
            Exit Function
        Case 368
            vla_step_text = "Ungroup rows 14 through 16. [line 689]" ' vla:4137
            Exit Function
        Case 369
            vla_step_text = "Put ""before-insert"" into cell A20. [line 691]" ' vla:4138
            Exit Function
        Case 370
            vla_step_text = "Insert 3 rows at row 20. [line 692]" ' vla:4139
            Exit Function
        Case 371
            vla_step_text = "Put ""before-delete"" into cell A40. [line 694]" ' vla:4140
            Exit Function
        Case 372
            vla_step_text = "Delete rows 38 through 39. [line 695]" ' vla:4141
            Exit Function
        Case 373
            vla_step_text = "Put ""marker-b"" into cell B50. [line 697]" ' vla:4142
            Exit Function
        Case 374
            vla_step_text = "Put ""marker-c"" into cell C50. [line 698]" ' vla:4143
            Exit Function
        Case 375
            vla_step_text = "Put ""marker-d"" into cell D50. [line 699]" ' vla:4144
            Exit Function
        Case 376
            vla_step_text = "Move column B before column D. [line 700]" ' vla:4145
            Exit Function
        Case 377
            vla_step_text = "Freeze the first 2 rows. [line 702]" ' vla:4146
            Exit Function
        Case 378
            vla_step_text = "Make cell A60 bold. [line 705]" ' vla:4147
            Exit Function
        Case 379
            vla_step_text = "Make cell A60 red. [line 706]" ' vla:4148
            Exit Function
        Case 380
            vla_step_text = "Copy formatting of A60:A60 to C60:C60. [line 707]" ' vla:4149
            Exit Function
        Case 381
            vla_step_text = "Make C62:C62 look like A60:A60. [line 708]" ' vla:4150
            Exit Function
        Case 382
            vla_step_text = "Put formula ""=5*2"" into cell B64. [line 710]" ' vla:4151
            Exit Function
        Case 383
            vla_step_text = "Copy formulas of B64 to D64. [line 711]" ' vla:4152
            Exit Function
        Case 384
            vla_step_text = "Set width of column F to 33. [line 713]" ' vla:4153
            Exit Function
        Case 385
            vla_step_text = "Copy column widths of F1:F1 to H1:H1. [line 714]" ' vla:4154
            Exit Function
        Case 386
            vla_step_text = "Put 1 into cell A68. [line 716]" ' vla:4155
            Exit Function
        Case 387
            vla_step_text = "Put 2 into cell A69. [line 717]" ' vla:4156
            Exit Function
        Case 388
            vla_step_text = "Put 3 into cell A70. [line 718]" ' vla:4157
            Exit Function
        Case 389
            vla_step_text = "Copy range A68:A70 to C68 transposed. [line 719]" ' vla:4158
            Exit Function
        Case 390
            vla_step_text = "Put ""cutme"" into cell A72. [line 721]" ' vla:4159
            Exit Function
        Case 391
            vla_step_text = "Move range A72:A72 to C72. [line 722]" ' vla:4160
            Exit Function
        Case 392
            vla_step_text = "Put ""rowdata"" into cell A74. [line 724]" ' vla:4161
            Exit Function
        Case 393
            vla_step_text = "Copy row 74 to row 76. [line 725]" ' vla:4162
            Exit Function
        Case 394
            vla_step_text = "Put ""clearme"" into cell A80. [line 728]" ' vla:4163
            Exit Function
        Case 395
            vla_step_text = "Make cell A80 bold. [line 729]" ' vla:4164
            Exit Function
        Case 396
            vla_step_text = "Make cell A80 red. [line 730]" ' vla:4165
            Exit Function
        Case 397
            vla_step_text = "Clear everything from A80:B80. [line 731]" ' vla:4166
            Exit Function
        Case 398
            vla_step_text = "Put ""m1"" into cell A84. [line 733]" ' vla:4167
            Exit Function
        Case 399
            vla_step_text = "Put ""m2"" into cell A85. [line 734]" ' vla:4168
            Exit Function
        Case 400
            vla_step_text = "Put ""m3"" into cell A86. [line 735]" ' vla:4169
            Exit Function
        Case 401
            vla_step_text = "Put ""m4"" into cell A87. [line 736]" ' vla:4170
            Exit Function
        Case 402
            vla_step_text = "Put ""m5"" into cell A88. [line 737]" ' vla:4171
            Exit Function
        Case 403
            vla_step_text = "Delete A84:A86 and shift cells up. [line 738]" ' vla:4172
            Exit Function
        Case 404
            vla_step_text = "Put ""x1"" into cell B90. [line 740]" ' vla:4173
            Exit Function
        Case 405
            vla_step_text = "Put ""x2"" into cell C90. [line 741]" ' vla:4174
            Exit Function
        Case 406
            vla_step_text = "Put ""x3"" into cell D90. [line 742]" ' vla:4175
            Exit Function
        Case 407
            vla_step_text = "Put ""rightdata"" into cell E90. [line 743]" ' vla:4176
            Exit Function
        Case 408
            vla_step_text = "Delete B90:D90 and shift cells left. [line 744]" ' vla:4177
            Exit Function
        Case 409
            vla_step_text = "Put 1 into cell A94. [line 746]" ' vla:4178
            Exit Function
        Case 410
            vla_step_text = "Put 2 into cell B94. [line 747]" ' vla:4179
            Exit Function
        Case 411
            vla_step_text = "Put 3 into cell A95. [line 748]" ' vla:4180
            Exit Function
        Case 412
            vla_step_text = "Put 5 into cell A96. [line 749]" ' vla:4181
            Exit Function
        Case 413
            vla_step_text = "Put 6 into cell B96. [line 750]" ' vla:4182
            Exit Function
        Case 414
            vla_step_text = "Delete blank rows in A94:B96. [line 751]" ' vla:4183
            Exit Function
        Case 415
            vla_step_text = "Put 1 into cell A150. [line 754]" ' vla:4184
            Exit Function
        Case 416
            vla_step_text = "Put 2 into cell A151. [line 755]" ' vla:4185
            Exit Function
        Case 417
            vla_step_text = "Put 3 into cell A152. [line 756]" ' vla:4186
            Exit Function
        Case 418
            vla_step_text = "Put 4 into cell A153. [line 757]" ' vla:4187
            Exit Function
        Case 419
            vla_step_text = "Band every other row of A150:D153 ""#D9D9D9"". [line 758]" ' vla:4188
            Exit Function
        Case 420
            vla_step_text = "Make row 165 a header row. [line 759]" ' vla:4189
            Exit Function
        Case 421
            vla_step_text = "Fill A170:A179 with a series starting at 1. [line 762]" ' vla:4190
            Exit Function
        Case 422
            vla_step_text = "Fill A180:A184 with a series starting at 1 with step 2. [line 763]" ' vla:4191
            Exit Function
        Case 423
            vla_step_text = "Fill A190:A194 with a growth series starting at 2. [line 764]" ' vla:4192
            Exit Function
        Case 424
            vla_step_text = "Fill A200:A203 with a growth series starting at 2 with step 3. [line 765]" ' vla:4193
            Exit Function
        Case 425
            vla_step_text = "Make cell Z500 bold. [line 777]" ' vla:4194
            Exit Function
        Case 426
            vla_step_text = "Clear everything from Z500. [line 778]" ' vla:4195
            Exit Function
        Case 427
            vla_step_text = "Remove trailing empty rows and columns. [line 779]" ' vla:4196
            Exit Function
        Case 428
            vla_step_text = "Try: [line 792]" ' vla:4197
            Exit Function
        Case 429
            vla_step_text = "Delete sheet GFormat. [line 793]" ' vla:4198
            Exit Function
        Case 430
            vla_step_text = "Work on sheet GFormat. [line 795]" ' vla:4199
            Exit Function
        Case 431
            vla_step_text = "Make range A1:C1 bold. [line 799]" ' vla:4200
            Exit Function
        Case 432
            vla_step_text = "Make range A2:C2 italic. [line 800]" ' vla:4201
            Exit Function
        Case 433
            vla_step_text = "Make range A3:C3 blue. [line 801]" ' vla:4202
            Exit Function
        Case 434
            vla_step_text = "Set font-color of range A4:C4 to ""#FF0000"". [line 802]" ' vla:4203
            Exit Function
        Case 435
            vla_step_text = "Set fill-color of range A5:C5 to ""#00FF00"". [line 803]" ' vla:4204
            Exit Function
        Case 436
            vla_step_text = "Make range A6:C6 yellow. [line 804]" ' vla:4205
            Exit Function
        Case 437
            vla_step_text = "Clear fill-color of range B6:C6. [line 805]" ' vla:4206
            Exit Function
        Case 438
            vla_step_text = "Set font size of range A7:C7 to 16. [line 806]" ' vla:4207
            Exit Function
        Case 439
            vla_step_text = "Make range A8:C8 bold. [line 809]" ' vla:4208
            Exit Function
        Case 440
            vla_step_text = "Make range B8:C8 not bold. [line 810]" ' vla:4209
            Exit Function
        Case 441
            vla_step_text = "Make range A9:C9 italic. [line 811]" ' vla:4210
            Exit Function
        Case 442
            vla_step_text = "Make cell C9 not italic. [line 812]" ' vla:4211
            Exit Function
        Case 443
            vla_step_text = "Underline range A10:C10. [line 813]" ' vla:4212
            Exit Function
        Case 444
            vla_step_text = "Strike through range A11:C11. [line 814]" ' vla:4213
            Exit Function
        Case 445
            vla_step_text = "Underline range A12:C12. [line 815]" ' vla:4214
            Exit Function
        Case 446
            vla_step_text = "Remove underline from range B12:C12. [line 816]" ' vla:4215
            Exit Function
        Case 447
            vla_step_text = "Strike through range A13:C13. [line 817]" ' vla:4216
            Exit Function
        Case 448
            vla_step_text = "Remove the strikethrough from cell C13. [line 818]" ' vla:4217
            Exit Function
        Case 449
            vla_step_text = "Set font of range A14:C14 to ""Courier New"". [line 819]" ' vla:4218
            Exit Function
        Case 450
            vla_step_text = "Align range A15:C15 to the top. [line 823]" ' vla:4219
            Exit Function
        Case 451
            vla_step_text = "Align range A16:C16 to the middle. [line 824]" ' vla:4220
            Exit Function
        Case 452
            vla_step_text = "Align range A17:C17 to the top. [line 825]" ' vla:4221
            Exit Function
        Case 453
            vla_step_text = "Align range B17:C17 to the bottom. [line 826]" ' vla:4222
            Exit Function
        Case 454
            vla_step_text = "Indent range A18:C18 by 2. [line 827]" ' vla:4223
            Exit Function
        Case 455
            vla_step_text = "Rotate text in range A19:C19 by 45 degrees. [line 828]" ' vla:4224
            Exit Function
        Case 456
            vla_step_text = "Add a border around range B21:D23. [line 834]" ' vla:4225
            Exit Function
        Case 457
            vla_step_text = "Add a bottom border to range B25:D25. [line 835]" ' vla:4226
            Exit Function
        Case 458
            vla_step_text = "Add borders to every cell in range B27:D29. [line 836]" ' vla:4227
            Exit Function
        Case 459
            vla_step_text = "Add borders to every cell in range B31:D33. [line 837]" ' vla:4228
            Exit Function
        Case 460
            vla_step_text = "Remove borders from range C31:D33. [line 838]" ' vla:4229
            Exit Function
        Case 461
            vla_step_text = "Set width of column F to 40. [line 846]" ' vla:4230
            Exit Function
        Case 462
            vla_step_text = "Put 1234.56 into cell F40. [line 847]" ' vla:4231
            Exit Function
        Case 463
            vla_step_text = "Format cell F40 as a number. [line 848]" ' vla:4232
            Exit Function
        Case 464
            vla_step_text = "Put 1234.56 into cell F41. [line 849]" ' vla:4233
            Exit Function
        Case 465
            vla_step_text = "Format cell F41 as a number with 3 decimals. [line 850]" ' vla:4234
            Exit Function
        Case 466
            vla_step_text = "Put 1234.56 into cell F42. [line 851]" ' vla:4235
            Exit Function
        Case 467
            vla_step_text = "Format cell F42 as a number with thousands separators. [line 852]" ' vla:4236
            Exit Function
        Case 468
            vla_step_text = "Put 1234.56 into cell F43. [line 853]" ' vla:4237
            Exit Function
        Case 469
            vla_step_text = "Format cell F43 as a number with 0 decimals and thousands separators. [line 854]" ' vla:4238
            Exit Function
        Case 470
            vla_step_text = "Put 1234.56 into cell F44. [line 855]" ' vla:4239
            Exit Function
        Case 471
            vla_step_text = "Format cell F44 as dollars. [line 856]" ' vla:4240
            Exit Function
        Case 472
            vla_step_text = "Put 1234.56 into cell F45. [line 857]" ' vla:4241
            Exit Function
        Case 473
            vla_step_text = "Format cell F45 as euros with 0 decimals. [line 858]" ' vla:4242
            Exit Function
        Case 474
            vla_step_text = "Put 1234.56 into cell F46. [line 859]" ' vla:4243
            Exit Function
        Case 475
            vla_step_text = "Format cell F46 as pounds. [line 860]" ' vla:4244
            Exit Function
        Case 476
            vla_step_text = "Put 1234.56 into cell F47. [line 861]" ' vla:4245
            Exit Function
        Case 477
            vla_step_text = "Format cell F47 as accounting in dollars. [line 862]" ' vla:4246
            Exit Function
        Case 478
            vla_step_text = "Put -1234.56 into cell F48. [line 863]" ' vla:4247
            Exit Function
        Case 479
            vla_step_text = "Format cell F48 as accounting in euros with 0 decimals. [line 864]" ' vla:4248
            Exit Function
        Case 480
            vla_step_text = "Put 0.125 into cell F49. [line 865]" ' vla:4249
            Exit Function
        Case 481
            vla_step_text = "Format range F49 as percent. [line 866]" ' vla:4250
            Exit Function
        Case 482
            vla_step_text = "Put 0.125 into cell F50. [line 867]" ' vla:4251
            Exit Function
        Case 483
            vla_step_text = "Format cell F50 as percent with 2 decimals. [line 868]" ' vla:4252
            Exit Function
        Case 484
            vla_step_text = "Put 46000 into cell F51. [line 869]" ' vla:4253
            Exit Function
        Case 485
            vla_step_text = "Format cell F51 as a short date. [line 870]" ' vla:4254
            Exit Function
        Case 486
            vla_step_text = "Put 46000 into cell F52. [line 871]" ' vla:4255
            Exit Function
        Case 487
            vla_step_text = "Format cell F52 as a long date. [line 872]" ' vla:4256
            Exit Function
        Case 488
            vla_step_text = "Put 46000 into cell F53. [line 873]" ' vla:4257
            Exit Function
        Case 489
            vla_step_text = "Format cell F53 as an ISO date. [line 874]" ' vla:4258
            Exit Function
        Case 490
            vla_step_text = "Put 0.5625 into cell F54. [line 875]" ' vla:4259
            Exit Function
        Case 491
            vla_step_text = "Format cell F54 as a time. [line 876]" ' vla:4260
            Exit Function
        Case 492
            vla_step_text = "Put 42 into cell F55. [line 877]" ' vla:4261
            Exit Function
        Case 493
            vla_step_text = "Format cell F55 as text for new entries. [line 878]" ' vla:4262
            Exit Function
        Case 494
            vla_step_text = "Put 1234.56 into cell F56. [line 879]" ' vla:4263
            Exit Function
        Case 495
            vla_step_text = "Format cell F56 as dollars. [line 880]" ' vla:4264
            Exit Function
        Case 496
            vla_step_text = "Format cell F56 as general. [line 881]" ' vla:4265
            Exit Function
        Case 497
            vla_step_text = "Put 42 into cell F57. [line 882]" ' vla:4266
            Exit Function
        Case 498
            vla_step_text = "Format cell F57 using ""00000"". [line 883]" ' vla:4267
            Exit Function
        Case 499
            vla_step_text = "Add a border colored ""#FF0000"" around range H40:J42. [line 889]" ' vla:4268
            Exit Function
        Case 500
            vla_step_text = "Add a bottom border colored green to range H44:J44. [line 890]" ' vla:4269
            Exit Function
        Case 501
            vla_step_text = "Add borders colored blue to every cell in range H46:J48. [line 891]" ' vla:4270
            Exit Function
        Case 502
            vla_step_text = "Try: [line 899]" ' vla:4271
            Exit Function
        Case 503
            vla_step_text = "Delete sheet GSortFilter. [line 900]" ' vla:4272
            Exit Function
        Case 504
            vla_step_text = "Work on sheet GSortFilter. [line 902]" ' vla:4273
            Exit Function
        Case 505
            vla_step_text = "Put ""Item"" into cell A1. [line 906]" ' vla:4274
            Exit Function
        Case 506
            vla_step_text = "Put ""Qty"" into cell B1. [line 907]" ' vla:4275
            Exit Function
        Case 507
            vla_step_text = "Put ""a"" into cell A2. [line 908]" ' vla:4276
            Exit Function
        Case 508
            vla_step_text = "Put 2 into cell B2. [line 909]" ' vla:4277
            Exit Function
        Case 509
            vla_step_text = "Put ""b"" into cell A3. [line 910]" ' vla:4278
            Exit Function
        Case 510
            vla_step_text = "Put 9 into cell B3. [line 911]" ' vla:4279
            Exit Function
        Case 511
            vla_step_text = "Put ""c"" into cell A4. [line 912]" ' vla:4280
            Exit Function
        Case 512
            vla_step_text = "Put 5 into cell B4. [line 913]" ' vla:4281
            Exit Function
        Case 513
            vla_step_text = "Sort this sheet by column B descending with a header row. [line 914]" ' vla:4282
            Exit Function
        Case 514
            vla_step_text = "Put ""Name"" into cell D1. [line 918]" ' vla:4283
            Exit Function
        Case 515
            vla_step_text = "Put ""Score"" into cell E1. [line 919]" ' vla:4284
            Exit Function
        Case 516
            vla_step_text = "Put ""p"" into cell D2. [line 920]" ' vla:4285
            Exit Function
        Case 517
            vla_step_text = "Put 3 into cell E2. [line 921]" ' vla:4286
            Exit Function
        Case 518
            vla_step_text = "Put ""q"" into cell D3. [line 922]" ' vla:4287
            Exit Function
        Case 519
            vla_step_text = "Put 1 into cell E3. [line 923]" ' vla:4288
            Exit Function
        Case 520
            vla_step_text = "Put ""r"" into cell D4. [line 924]" ' vla:4289
            Exit Function
        Case 521
            vla_step_text = "Put 2 into cell E4. [line 925]" ' vla:4290
            Exit Function
        Case 522
            vla_step_text = "Sort range D1:E4 by column E with a header row. [line 926]" ' vla:4291
            Exit Function
        Case 523
            vla_step_text = "Put ""x"" into cell G2. [line 929]" ' vla:4292
            Exit Function
        Case 524
            vla_step_text = "Put 5 into cell H2. [line 930]" ' vla:4293
            Exit Function
        Case 525
            vla_step_text = "Put ""y"" into cell G3. [line 931]" ' vla:4294
            Exit Function
        Case 526
            vla_step_text = "Put 9 into cell H3. [line 932]" ' vla:4295
            Exit Function
        Case 527
            vla_step_text = "Put ""z"" into cell G4. [line 933]" ' vla:4296
            Exit Function
        Case 528
            vla_step_text = "Put 7 into cell H4. [line 934]" ' vla:4297
            Exit Function
        Case 529
            vla_step_text = "Sort range G2:H4 by column H descending without a header row. [line 935]" ' vla:4298
            Exit Function
        Case 530
            vla_step_text = "Put ""Region"" into cell J1. [line 939]" ' vla:4299
            Exit Function
        Case 531
            vla_step_text = "Put ""Amount"" into cell K1. [line 940]" ' vla:4300
            Exit Function
        Case 532
            vla_step_text = "Put ""Tag"" into cell L1. [line 941]" ' vla:4301
            Exit Function
        Case 533
            vla_step_text = "Put ""West"" into cell J2. [line 942]" ' vla:4302
            Exit Function
        Case 534
            vla_step_text = "Put 100 into cell K2. [line 943]" ' vla:4303
            Exit Function
        Case 535
            vla_step_text = "Put ""a"" into cell L2. [line 944]" ' vla:4304
            Exit Function
        Case 536
            vla_step_text = "Put ""East"" into cell J3. [line 945]" ' vla:4305
            Exit Function
        Case 537
            vla_step_text = "Put 50 into cell K3. [line 946]" ' vla:4306
            Exit Function
        Case 538
            vla_step_text = "Put ""b"" into cell L3. [line 947]" ' vla:4307
            Exit Function
        Case 539
            vla_step_text = "Put ""West"" into cell J4. [line 948]" ' vla:4308
            Exit Function
        Case 540
            vla_step_text = "Put 300 into cell K4. [line 949]" ' vla:4309
            Exit Function
        Case 541
            vla_step_text = "Put ""c"" into cell L4. [line 950]" ' vla:4310
            Exit Function
        Case 542
            vla_step_text = "Put ""East"" into cell J5. [line 951]" ' vla:4311
            Exit Function
        Case 543
            vla_step_text = "Put 250 into cell K5. [line 952]" ' vla:4312
            Exit Function
        Case 544
            vla_step_text = "Put ""d"" into cell L5. [line 953]" ' vla:4313
            Exit Function
        Case 545
            vla_step_text = "Sort range J1:L5 by column J ascending then by column K descending with a header row. [line 954]" ' vla:4314
            Exit Function
        Case 546
            vla_step_text = "Put 2 into cell N2. [line 957]" ' vla:4315
            Exit Function
        Case 547
            vla_step_text = "Put 9 into cell O2. [line 958]" ' vla:4316
            Exit Function
        Case 548
            vla_step_text = "Put 1 into cell N3. [line 959]" ' vla:4317
            Exit Function
        Case 549
            vla_step_text = "Put 8 into cell O3. [line 960]" ' vla:4318
            Exit Function
        Case 550
            vla_step_text = "Put 2 into cell N4. [line 961]" ' vla:4319
            Exit Function
        Case 551
            vla_step_text = "Put 7 into cell O4. [line 962]" ' vla:4320
            Exit Function
        Case 552
            vla_step_text = "Sort range N2:O4 by column N then by column O without a header row. [line 963]" ' vla:4321
            Exit Function
        Case 553
            vla_step_text = "Put ""Region"" into cell A40. [line 972]" ' vla:4322
            Exit Function
        Case 554
            vla_step_text = "Put ""Amount"" into cell B40. [line 973]" ' vla:4323
            Exit Function
        Case 555
            vla_step_text = "Put ""West"" into cell A41. [line 974]" ' vla:4324
            Exit Function
        Case 556
            vla_step_text = "Put 1 into cell B41. [line 975]" ' vla:4325
            Exit Function
        Case 557
            vla_step_text = "Put ""East"" into cell A42. [line 976]" ' vla:4326
            Exit Function
        Case 558
            vla_step_text = "Put 2 into cell B42. [line 977]" ' vla:4327
            Exit Function
        Case 559
            vla_step_text = "Put ""West"" into cell A43. [line 978]" ' vla:4328
            Exit Function
        Case 560
            vla_step_text = "Put 3 into cell B43. [line 979]" ' vla:4329
            Exit Function
        Case 561
            vla_step_text = "Filter range A40:B43 to show rows where column A is ""West"". [line 980]" ' vla:4330
            Exit Function
        Case 562
            vla_step_text = "Copy only the visible cells of range A40:B43 to D46. [line 981]" ' vla:4331
            Exit Function
        Case 563
            vla_step_text = "Remove the filters. [line 982]" ' vla:4332
            Exit Function
        Case 564
            vla_step_text = "Try: [line 988]" ' vla:4333
            Exit Function
        Case 565
            vla_step_text = "Filter range A40:B43 to show rows where column B is greater than ""abc"". [line 989]" ' vla:4334
            Exit Function
        Case 566
            vla_step_text = "Put ""Region"" into cell A20. [line 994]" ' vla:4335
            Exit Function
        Case 567
            vla_step_text = "Put ""Amount"" into cell B20. [line 995]" ' vla:4336
            Exit Function
        Case 568
            vla_step_text = "Put ""Code"" into cell C20. [line 996]" ' vla:4337
            Exit Function
        Case 569
            vla_step_text = "Put ""West"" into cell A21. [line 997]" ' vla:4338
            Exit Function
        Case 570
            vla_step_text = "Put 100 into cell B21. [line 998]" ' vla:4339
            Exit Function
        Case 571
            vla_step_text = "Put ""5*3"" into cell C21. [line 999]" ' vla:4340
            Exit Function
        Case 572
            vla_step_text = "Put ""East"" into cell A22. [line 1000]" ' vla:4341
            Exit Function
        Case 573
            vla_step_text = "Put 250 into cell B22. [line 1001]" ' vla:4342
            Exit Function
        Case 574
            vla_step_text = "Put ""53"" into cell C22. [line 1002]" ' vla:4343
            Exit Function
        Case 575
            vla_step_text = "Put ""Western"" into cell A23. [line 1003]" ' vla:4344
            Exit Function
        Case 576
            vla_step_text = "Put 100 into cell B23. [line 1004]" ' vla:4345
            Exit Function
        Case 577
            vla_step_text = "Put ""x"" into cell C23. [line 1005]" ' vla:4346
            Exit Function
        Case 578
            vla_step_text = "Put ""West"" into cell A24. [line 1006]" ' vla:4347
            Exit Function
        Case 579
            vla_step_text = "Put 50 into cell B24. [line 1007]" ' vla:4348
            Exit Function
        Case 580
            vla_step_text = "Put ""y"" into cell C24. [line 1008]" ' vla:4349
            Exit Function
        Case 581
            vla_step_text = "Put ""east"" into cell A25. [line 1009]" ' vla:4350
            Exit Function
        Case 582
            vla_step_text = "Put 300 into cell B25. [line 1010]" ' vla:4351
            Exit Function
        Case 583
            vla_step_text = "Put ""z"" into cell C25. [line 1011]" ' vla:4352
            Exit Function
        Case 584
            vla_step_text = "Put ""North"" into cell A26. [line 1012]" ' vla:4353
            Exit Function
        Case 585
            vla_step_text = "Put 175 into cell B26. [line 1013]" ' vla:4354
            Exit Function
        Case 586
            vla_step_text = "Put ""w"" into cell C26. [line 1014]" ' vla:4355
            Exit Function
        Case 587
            vla_step_text = "Put ""East"" into cell A27. [line 1015]" ' vla:4356
            Exit Function
        Case 588
            vla_step_text = "Put 75 into cell B27. [line 1016]" ' vla:4357
            Exit Function
        Case 589
            vla_step_text = "Put ""v"" into cell C27. [line 1017]" ' vla:4358
            Exit Function
        Case 590
            vla_step_text = "Format range B21:B27 as dollars. [line 1018]" ' vla:4359
            Exit Function
        Case 591
            vla_step_text = "Filter range A20:C27 to show rows where column A is ""West"". [line 1020]" ' vla:4360
            Exit Function
        Case 592
            vla_step_text = "Copy only the visible cells of range A20:C27 to E30. [line 1021]" ' vla:4361
            Exit Function
        Case 593
            vla_step_text = "Clear the filter conditions. [line 1022]" ' vla:4362
            Exit Function
        Case 594
            vla_step_text = "Filter range A20:C27 to show rows where column B is 100. [line 1024]" ' vla:4363
            Exit Function
        Case 595
            vla_step_text = "Copy only the visible cells of range A20:C27 to I30. [line 1025]" ' vla:4364
            Exit Function
        Case 596
            vla_step_text = "Clear the filter conditions. [line 1026]" ' vla:4365
            Exit Function
        Case 597
            vla_step_text = "Filter range A20:C27 to show rows where column C contains ""*"". [line 1028]" ' vla:4366
            Exit Function
        Case 598
            vla_step_text = "Copy only the visible cells of range A20:C27 to M30. [line 1029]" ' vla:4367
            Exit Function
        Case 599
            vla_step_text = "Clear the filter conditions. [line 1030]" ' vla:4368
            Exit Function
        Case 600
            vla_step_text = "Filter range A20:C27 to show rows where column B is greater than 150. [line 1033]" ' vla:4369
            Exit Function
        Case 601
            vla_step_text = "Filter range A20:C27 to show rows where column A is ""East"". [line 1034]" ' vla:4370
            Exit Function
        Case 602
            vla_step_text = "Copy only the visible cells of range A20:C27 to Q30. [line 1035]" ' vla:4371
            Exit Function
        Case 603
            vla_step_text = "Clear the filter conditions. [line 1036]" ' vla:4372
            Exit Function
        Case 604
            vla_step_text = "Filter range A20:C27 to show rows where column B is less than 99.5. [line 1041]" ' vla:4373
            Exit Function
        Case 605
            vla_step_text = "Copy only the visible cells of range A20:C27 to U30. [line 1042]" ' vla:4374
            Exit Function
        Case 606
            vla_step_text = "Add filters to range A20:C27. [line 1049]" ' vla:4375
            Exit Function
        Case 607
            vla_step_text = "Clear the filter conditions. [line 1050]" ' vla:4376
            Exit Function
        Case 608
            vla_step_text = "Try: [line 1051]" ' vla:4377
            Exit Function
        Case 609
            vla_step_text = "Add filters to range A40:B43. [line 1052]" ' vla:4378
            Exit Function
        Case 610
            vla_step_text = "Try: [line 1060]" ' vla:4379
            Exit Function
        Case 611
            vla_step_text = "Delete sheet GText. [line 1061]" ' vla:4380
            Exit Function
        Case 612
            vla_step_text = "Work on sheet GText. [line 1063]" ' vla:4381
            Exit Function
        Case 613
            vla_step_text = "Put ""widget"" into cell A1. [line 1067]" ' vla:4382
            Exit Function
        Case 614
            vla_step_text = "Put 42 into cell A2. [line 1068]" ' vla:4383
            Exit Function
        Case 615
            vla_step_text = "Put formula ""=CHAR(97)&CHAR(98)"" into cell A3. [line 1069]" ' vla:4384
            Exit Function
        Case 616
            vla_step_text = "Put ""'true"" into cell A4. [line 1070]" ' vla:4385
            Exit Function
        Case 617
            vla_step_text = "Put ""'=abc"" into cell A5. [line 1071]" ' vla:4386
            Exit Function
        Case 618
            vla_step_text = "Make range A1:A5 upper case. [line 1072]" ' vla:4387
            Exit Function
        Case 619
            vla_step_text = "Put ""HELLO World"" into cell B1. [line 1075]" ' vla:4388
            Exit Function
        Case 620
            vla_step_text = "Put ""MiXeD"" into cell B2. [line 1076]" ' vla:4389
            Exit Function
        Case 621
            vla_step_text = "Make column B lowercase. [line 1077]" ' vla:4390
            Exit Function
        Case 622
            vla_step_text = "Put ""don't stop"" into cell C1. [line 1080]" ' vla:4391
            Exit Function
        Case 623
            vla_step_text = "Put ""3rd quarter"" into cell C2. [line 1081]" ' vla:4392
            Exit Function
        Case 624
            vla_step_text = "Put ""o'neil"" into cell C3. [line 1082]" ' vla:4393
            Exit Function
        Case 625
            vla_step_text = "Capitalize each word in range C1:C3 after any space. [line 1083]" ' vla:4394
            Exit Function
        Case 626
            vla_step_text = "Put ""don't stop"" into cell D1. [line 1084]" ' vla:4395
            Exit Function
        Case 627
            vla_step_text = "Put ""3rd quarter"" into cell D2. [line 1085]" ' vla:4396
            Exit Function
        Case 628
            vla_step_text = "Put ""o'neil"" into cell D3. [line 1086]" ' vla:4397
            Exit Function
        Case 629
            vla_step_text = "Capitalize each word in column D after any non-letter. [line 1087]" ' vla:4398
            Exit Function
        Case 630
            vla_step_text = "Put ""  a   b  "" into cell E1. [line 1092]" ' vla:4399
            Exit Function
        Case 631
            vla_step_text = "Put ""'  00123 "" into cell E2. [line 1093]" ' vla:4400
            Exit Function
        Case 632
            vla_step_text = "Put ""   "" into cell E3. [line 1094]" ' vla:4401
            Exit Function
        Case 633
            vla_step_text = "Put formula ""=UNICHAR(160)&CHAR(120)&UNICHAR(160)&UNICHAR(160)&CHAR(121)"" into cell E4. [line 1095]" ' vla:4402
            Exit Function
        Case 634
            vla_step_text = "Convert range E4:E4 to values. [line 1096]" ' vla:4403
            Exit Function
        Case 635
            vla_step_text = "Remove extra spaces from range E1:E4. [line 1097]" ' vla:4404
            Exit Function
        Case 636
            vla_step_text = "Put formula ""=CHAR(97)&CHAR(10)&CHAR(98)"" into cell F1. [line 1100]" ' vla:4405
            Exit Function
        Case 637
            vla_step_text = "Put formula ""=CHAR(99)&CHAR(9)&CHAR(100)"" into cell F2. [line 1101]" ' vla:4406
            Exit Function
        Case 638
            vla_step_text = "Convert range F1:F2 to values. [line 1102]" ' vla:4407
            Exit Function
        Case 639
            vla_step_text = "Remove non-printing characters from range F1:F2. [line 1103]" ' vla:4408
            Exit Function
        Case 640
            vla_step_text = "Put ""Widget"" into cell G1. [line 1108]" ' vla:4409
            Exit Function
        Case 641
            vla_step_text = "Put ""banana"" into cell G2. [line 1109]" ' vla:4410
            Exit Function
        Case 642
            vla_step_text = "Set gtext-row to row of ""Widget"" in column G. [line 1110]" ' vla:4411
            Exit Function
        Case 643
            vla_step_text = "Replace ""AN"" with ""xy"" in range G2:G2. [line 1111]" ' vla:4412
            Exit Function
        Case 644
            vla_step_text = "Format range I1:I20 as text for new entries. [line 1115]" ' vla:4413
            Exit Function
        Case 645
            vla_step_text = "Set gtext-code to ""INV-ab-Cd"". [line 1116]" ' vla:4414
            Exit Function
        Case 646
            vla_step_text = "Set gtext-part to the text before ""-"" in gtext-code. [line 1117]" ' vla:4415
            Exit Function
        Case 647
            vla_step_text = "Put gtext-part into cell I1. [line 1118]" ' vla:4416
            Exit Function
        Case 648
            vla_step_text = "Set gtext-part to the text after ""-"" in gtext-code. [line 1119]" ' vla:4417
            Exit Function
        Case 649
            vla_step_text = "Put gtext-part into cell I2. [line 1120]" ' vla:4418
            Exit Function
        Case 650
            vla_step_text = "Set gtext-part to the text after the last ""-"" in gtext-code. [line 1121]" ' vla:4419
            Exit Function
        Case 651
            vla_step_text = "Put gtext-part into cell I3. [line 1122]" ' vla:4420
            Exit Function
        Case 652
            vla_step_text = "Set gtext-part to the text before ""c"" in gtext-code. [line 1125]" ' vla:4421
            Exit Function
        Case 653
            vla_step_text = "Put gtext-part into cell I4. [line 1126]" ' vla:4422
            Exit Function
        Case 654
            vla_step_text = "Set gtext-part to the first 3 characters of gtext-code. [line 1127]" ' vla:4423
            Exit Function
        Case 655
            vla_step_text = "Put gtext-part into cell I5. [line 1128]" ' vla:4424
            Exit Function
        Case 656
            vla_step_text = "Set gtext-part to the last 2 characters of gtext-code. [line 1129]" ' vla:4425
            Exit Function
        Case 657
            vla_step_text = "Put gtext-part into cell I6. [line 1130]" ' vla:4426
            Exit Function
        Case 658
            vla_step_text = "Set gtext-part to 42 padded on the left with ""0"" to 5 characters. [line 1132]" ' vla:4427
            Exit Function
        Case 659
            vla_step_text = "Put gtext-part into cell I7. [line 1133]" ' vla:4428
            Exit Function
        Case 660
            vla_step_text = "Set gtext-part to ""ab"" padded on the right with ""."" to 4 characters. [line 1134]" ' vla:4429
            Exit Function
        Case 661
            vla_step_text = "Put gtext-part into cell I8. [line 1135]" ' vla:4430
            Exit Function
        Case 662
            vla_step_text = "Put ""x"" into cell J1. [line 1138]" ' vla:4431
            Exit Function
        Case 663
            vla_step_text = "Put 7 into cell J3. [line 1139]" ' vla:4432
            Exit Function
        Case 664
            vla_step_text = "Put ""y"" into cell J4. [line 1140]" ' vla:4433
            Exit Function
        Case 665
            vla_step_text = "Set gtext-list to range J1:J4 as one list. [line 1141]" ' vla:4434
            Exit Function
        Case 666
            vla_step_text = "Put gtext-list into cell I9. [line 1142]" ' vla:4435
            Exit Function
        Case 667
            vla_step_text = "Set gtext-list to range J1:J4 as one list separated by ""; "". [line 1143]" ' vla:4436
            Exit Function
        Case 668
            vla_step_text = "Put gtext-list into cell I10. [line 1144]" ' vla:4437
            Exit Function
        Case 669
            vla_step_text = "Set gtext-list to column J as one list. [line 1145]" ' vla:4438
            Exit Function
        Case 670
            vla_step_text = "Put gtext-list into cell I11. [line 1146]" ' vla:4439
            Exit Function
        Case 671
            vla_step_text = "Set gtext-part to ""  a   b "" with extra spaces removed. [line 1150]" ' vla:4440
            Exit Function
        Case 672
            vla_step_text = "Put gtext-part into cell I12. [line 1151]" ' vla:4441
            Exit Function
        Case 673
            vla_step_text = "Set gtext-part to ""o'neil"" with each word capitalized after any non-letter. [line 1152]" ' vla:4442
            Exit Function
        Case 674
            vla_step_text = "Put gtext-part into cell I13. [line 1153]" ' vla:4443
            Exit Function
        Case 675
            vla_step_text = "Set gtext-part to ""don't stop"" with each word capitalized after any space. [line 1154]" ' vla:4444
            Exit Function
        Case 676
            vla_step_text = "Put gtext-part into cell I14. [line 1155]" ' vla:4445
            Exit Function
        Case 677
            vla_step_text = "Put formula ""=CHAR(112)&CHAR(10)&CHAR(113)"" into cell K1. [line 1156]" ' vla:4446
            Exit Function
        Case 678
            vla_step_text = "Convert range K1:K1 to values. [line 1157]" ' vla:4447
            Exit Function
        Case 679
            vla_step_text = "Set gtext-part to cell K1 with non-printing characters removed. [line 1158]" ' vla:4448
            Exit Function
        Case 680
            vla_step_text = "Put gtext-part into cell I15. [line 1159]" ' vla:4449
            Exit Function
        Case 681
            vla_step_text = "Try: [line 1164]" ' vla:4450
            Exit Function
        Case 682
            vla_step_text = "Set gtext-part to the text before ""#"" in gtext-code. [line 1165]" ' vla:4451
            Exit Function
        Case 683
            vla_step_text = "Put gtext-part into cell I16. [line 1167]" ' vla:4452
            Exit Function
        Case 684
            vla_step_text = "Put ""Nan"" into cell L1. [line 1173]" ' vla:4453
            Exit Function
        Case 685
            vla_step_text = "Put formula ""=TAN(0)"" into cell L2. [line 1174]" ' vla:4454
            Exit Function
        Case 686
            vla_step_text = "Replace ""an"" with ""xy"" in range L1:L2. [line 1175]" ' vla:4455
            Exit Function
        Case 687
            vla_step_text = "Put ""Q"" into cell L3. [line 1176]" ' vla:4456
            Exit Function
        Case 688
            vla_step_text = "Replace ""Q"" with ""=1+1"" in range L3:L3. [line 1177]" ' vla:4457
            Exit Function
        Case 689
            vla_step_text = "Put ""N/A"" into cell L4. [line 1178]" ' vla:4458
            Exit Function
        Case 690
            vla_step_text = "Replace ""N/A"" with 0 in range L4:L4. [line 1179]" ' vla:4459
            Exit Function
        Case 691
            vla_step_text = "Put formula ""=ABS(-2)"" into cell L5. [line 1182]" ' vla:4460
            Exit Function
        Case 692
            vla_step_text = "Put ""ABS"" into cell L6. [line 1183]" ' vla:4461
            Exit Function
        Case 693
            vla_step_text = "Replace ""ABS"" with ""SIGN"" in the formulas of range L5:L6. [line 1184]" ' vla:4462
            Exit Function
        Case 694
            vla_step_text = "Put ""zqz"" into cell L7. [line 1187]" ' vla:4463
            Exit Function
        Case 695
            vla_step_text = "Replace ""zqz"" with ""done"" on this sheet. [line 1188]" ' vla:4464
            Exit Function
        Case 696
            vla_step_text = "Put formula ""=ABS(-3)"" into cell L8. [line 1189]" ' vla:4465
            Exit Function
        Case 697
            vla_step_text = "Replace ""-3"" with ""-4"" in the formulas on this sheet. [line 1190]" ' vla:4466
            Exit Function
        Case 698
            vla_step_text = "Put ""x,0042,7"" into cell N1. [line 1195]" ' vla:4467
            Exit Function
        Case 699
            vla_step_text = "Split column N by "","" as text. [line 1196]" ' vla:4468
            Exit Function
        Case 700
            vla_step_text = "Put ""y,0042,7,-3.5"" into cell R1. [line 1197]" ' vla:4469
            Exit Function
        Case 701
            vla_step_text = "Split column R by "","" reading numbers. [line 1198]" ' vla:4470
            Exit Function
        Case 702
            vla_step_text = "Put ""p,q"" into cell W1. [line 1199]" ' vla:4471
            Exit Function
        Case 703
            vla_step_text = "Put ""keep"" into cell X1. [line 1200]" ' vla:4472
            Exit Function
        Case 704
            vla_step_text = "Try: [line 1201]" ' vla:4473
            Exit Function
        Case 705
            vla_step_text = "Split column W by "","" as text. [line 1202]" ' vla:4474
            Exit Function
        Case 706
            vla_step_text = "Put ""apple"" into cell Z1. [line 1205]" ' vla:4475
            Exit Function
        Case 707
            vla_step_text = "Put ""Banana"" into cell AA1. [line 1206]" ' vla:4476
            Exit Function
        Case 708
            vla_step_text = "Put ""cherry"" into cell Z2. [line 1207]" ' vla:4477
            Exit Function
        Case 709
            vla_step_text = "Put ""banana split"" into cell AA2. [line 1208]" ' vla:4478
            Exit Function
        Case 710
            vla_step_text = "Put 7 into cell Z3. [line 1209]" ' vla:4479
            Exit Function
        Case 711
            vla_step_text = "Put ""BANANA"" into cell AA3. [line 1210]" ' vla:4480
            Exit Function
        Case 712
            vla_step_text = "Set gtext-found to the row of the first cell in range Z1:AA3 containing ""nan"". [line 1211]" ' vla:4481
            Exit Function
        Case 713
            vla_step_text = "Put gtext-found into cell L9. [line 1212]" ' vla:4482
            Exit Function
        Case 714
            vla_step_text = "Set gtext-found to the column of the first cell in range Z1:AA3 containing ""nan"". [line 1213]" ' vla:4483
            Exit Function
        Case 715
            vla_step_text = "Put gtext-found into cell L10. [line 1214]" ' vla:4484
            Exit Function
        Case 716
            vla_step_text = "Set gtext-found to the row of the first cell in column Z containing ""err"". [line 1215]" ' vla:4485
            Exit Function
        Case 717
            vla_step_text = "Put gtext-found into cell L11. [line 1216]" ' vla:4486
            Exit Function
        Case 718
            vla_step_text = "Set gtext-found to how many cells in range Z1:AA3 contain ""banana"". [line 1217]" ' vla:4487
            Exit Function
        Case 719
            vla_step_text = "Put gtext-found into cell L12. [line 1218]" ' vla:4488
            Exit Function
        Case 720
            vla_step_text = "Set gtext-found to how many cells in column Z contain ""zzz"". [line 1219]" ' vla:4489
            Exit Function
        Case 721
            vla_step_text = "Put gtext-found into cell L13. [line 1220]" ' vla:4490
            Exit Function
        Case 722
            vla_step_text = "Set gtext-found to the row of the first cell in range Z1:AA3 containing ""kiwi"". [line 1221]" ' vla:4491
            Exit Function
        Case 723
            vla_step_text = "Put gtext-found into cell L14. [line 1222]" ' vla:4492
            Exit Function
        Case 724
            vla_step_text = "If range Z1:AA3 contains ""split"", put ""yes"" into cell L15. [line 1223]" ' vla:4493
            Exit Function
        Case 725
            vla_step_text = "If column Z does not contain ""kiwi"", put ""no kiwi"" into cell L16. [line 1224]" ' vla:4494
            Exit Function
        Case 726
            vla_step_text = "If cell AA1 contains ""NAN"", put ""contains"" into cell L17. [line 1238]" ' vla:4495
            Exit Function
        Case 727
            vla_step_text = "If cell AA1 starts with ""ban"", put ""starts"" into cell L18. [line 1239]" ' vla:4496
            Exit Function
        Case 728
            vla_step_text = "Put ""kept"" into cell L19. [line 1240]" ' vla:4497
            Exit Function
        Case 729
            vla_step_text = "If cell AA1 does not contain ""nan"", put ""wrong"" into cell L19. [line 1241]" ' vla:4498
            Exit Function
        Case 730
            vla_step_text = "If cell AA1 does not contain ""kiwi"", put ""lacks"" into cell L20. [line 1242]" ' vla:4499
            Exit Function
        Case 731
            vla_step_text = "Put ""kept"" into cell L21. [line 1243]" ' vla:4500
            Exit Function
        Case 732
            vla_step_text = "If cell AA1 contains ""kiwi"", put ""wrong"" into cell L21. [line 1244]" ' vla:4501
            Exit Function
        Case 733
            vla_step_text = "Put ""kept"" into cell L22. [line 1245]" ' vla:4502
            Exit Function
        Case 734
            vla_step_text = "If cell AA1 starts with ""nan"", put ""wrong"" into cell L22. [line 1246]" ' vla:4503
            Exit Function
        Case 735
            vla_step_text = "Check-text-conditions. [line 1248]" ' vla:4504
            Exit Function
        Case 736
            vla_step_text = "Work on sheet GFormula. [line 1259]" ' vla:4505
            Exit Function
        Case 737
            vla_step_text = "Clear everything from A1:H10. [line 1260]" ' vla:4506
            Exit Function
        Case 738
            vla_step_text = "Put 2 into cell B2. [line 1261]" ' vla:4507
            Exit Function
        Case 739
            vla_step_text = "Put 3 into cell B3. [line 1262]" ' vla:4508
            Exit Function
        Case 740
            vla_step_text = "Put 4 into cell B4. [line 1263]" ' vla:4509
            Exit Function
        Case 741
            vla_step_text = "Put 10 into cell C2. [line 1264]" ' vla:4510
            Exit Function
        Case 742
            vla_step_text = "Put 20 into cell C3. [line 1265]" ' vla:4511
            Exit Function
        Case 743
            vla_step_text = "Put 30 into cell C4. [line 1266]" ' vla:4512
            Exit Function
        Case 744
            vla_step_text = "Put formula ""=B2*C2"" into range D2:D4. [line 1267]" ' vla:4513
            Exit Function
        Case 745
            vla_step_text = "Set gf-last to last filled row of column B. [line 1268]" ' vla:4514
            Exit Function
        Case 746
            vla_step_text = "Put formula ""=B2+C2"" into rows 2 to gf-last of column E. [line 1269]" ' vla:4515
            Exit Function
        Case 747
            vla_step_text = "Put ""head"" into cell F1. [line 1270]" ' vla:4516
            Exit Function
        Case 748
            vla_step_text = "Put formula ""=B2-C2"" into rows 2 through 1 of column F. [line 1271]" ' vla:4517
            Exit Function
        Case 749
            vla_step_text = "Set gf-top to largest of range D2:D4. [line 1272]" ' vla:4518
            Exit Function
        Case 750
            vla_step_text = "Put gf-top into cell H1. [line 1273]" ' vla:4519
            Exit Function
        Case 751
            vla_step_text = "Set gf-bottom to smallest of range D2:D4. [line 1274]" ' vla:4520
            Exit Function
        Case 752
            vla_step_text = "Put gf-bottom into cell H2. [line 1275]" ' vla:4521
            Exit Function
        Case 753
            vla_step_text = "Put average of range B2:B4 into cell H3. [line 1276]" ' vla:4522
            Exit Function
        Case 754
            vla_step_text = "Put largest of range C2:C4 into cell H4. [line 1277]" ' vla:4523
            Exit Function
        Case 755
            vla_step_text = "Put smallest of range C2:C4 in cell H5. [line 1278]" ' vla:4524
            Exit Function
        Case 756
            vla_step_text = "Put formula ""="""""""""" into cell B6. [line 1279]" ' vla:4525
            Exit Function
        Case 757
            vla_step_text = "Set gf-empty to count of empty cells in range B2:B7. [line 1280]" ' vla:4526
            Exit Function
        Case 758
            vla_step_text = "Put gf-empty into cell H6. [line 1281]" ' vla:4527
            Exit Function
        Case 759
            vla_step_text = "Set gf-filled to count of filled cells in range B2:B7. [line 1282]" ' vla:4528
            Exit Function
        Case 760
            vla_step_text = "Put gf-filled into cell H7. [line 1283]" ' vla:4529
            Exit Function
        Case 761
            vla_step_text = "Check-formulas. [line 1285]" ' vla:4530
            Exit Function
        Case 762
            vla_step_text = "Work on sheet GStats. [line 1300]" ' vla:4531
            Exit Function
        Case 763
            vla_step_text = "Clear everything from A1:H10. [line 1301]" ' vla:4532
            Exit Function
        Case 764
            vla_step_text = "Put 2 into cell B2. [line 1302]" ' vla:4533
            Exit Function
        Case 765
            vla_step_text = "Put 4 into cell B3. [line 1303]" ' vla:4534
            Exit Function
        Case 766
            vla_step_text = "Put 6 into cell B4. [line 1304]" ' vla:4535
            Exit Function
        Case 767
            vla_step_text = "Put 6 into cell B5. [line 1305]" ' vla:4536
            Exit Function
        Case 768
            vla_step_text = "Put 7 into cell B6. [line 1306]" ' vla:4537
            Exit Function
        Case 769
            vla_step_text = "Set gs-middle to median of range B2:B6. [line 1307]" ' vla:4538
            Exit Function
        Case 770
            vla_step_text = "Put gs-middle into cell H1. [line 1308]" ' vla:4539
            Exit Function
        Case 771
            vla_step_text = "Set gs-spread to standard deviation of range B2:B6 as a sample. [line 1309]" ' vla:4540
            Exit Function
        Case 772
            vla_step_text = "Put gs-spread into cell H2. [line 1310]" ' vla:4541
            Exit Function
        Case 773
            vla_step_text = "Put median of range B2:B5 into cell H3. [line 1311]" ' vla:4542
            Exit Function
        Case 774
            vla_step_text = "Put standard deviation of range B2:B6 as the population into cell H4. [line 1312]" ' vla:4543
            Exit Function
        Case 775
            vla_step_text = "Put ""West"" into cell D2. [line 1313]" ' vla:4544
            Exit Function
        Case 776
            vla_step_text = "Put ""East"" into cell D3. [line 1314]" ' vla:4545
            Exit Function
        Case 777
            vla_step_text = "Put ""West"" into cell D4. [line 1315]" ' vla:4546
            Exit Function
        Case 778
            vla_step_text = "Put ""West"" into cell D5. [line 1316]" ' vla:4547
            Exit Function
        Case 779
            vla_step_text = "Put 10 into cell E2. [line 1317]" ' vla:4548
            Exit Function
        Case 780
            vla_step_text = "Put 20 into cell E3. [line 1318]" ' vla:4549
            Exit Function
        Case 781
            vla_step_text = "Put 30 into cell E5. [line 1319]" ' vla:4550
            Exit Function
        Case 782
            vla_step_text = "Set gs-west to average of range E2:E5 where range D2:D5 matches ""West"". [line 1320]" ' vla:4551
            Exit Function
        Case 783
            vla_step_text = "Put gs-west into cell H5. [line 1321]" ' vla:4552
            Exit Function
        Case 784
            vla_step_text = "Put ""kept"" into cell H6. [line 1322]" ' vla:4553
            Exit Function
        Case 785
            vla_step_text = "Try: [line 1323]" ' vla:4554
            Exit Function
        Case 786
            vla_step_text = "Set gs-none to median of range G1:G3. [line 1324]" ' vla:4555
            Exit Function
        Case 787
            vla_step_text = "Put ""no stop"" into cell H6. [line 1325]" ' vla:4556
            Exit Function
        Case 788
            vla_step_text = "Put the problem into cell H7. [line 1328]" ' vla:4557
            Exit Function
        Case 789
            vla_step_text = "Check-statistics. [line 1330]" ' vla:4558
            Exit Function
        Case 790
            vla_step_text = "Work on sheet GCalc. [line 1343]" ' vla:4559
            Exit Function
        Case 791
            vla_step_text = "Clear everything from A1:D10. [line 1344]" ' vla:4560
            Exit Function
        Case 792
            vla_step_text = "Put 2 into cell A1. [line 1345]" ' vla:4561
            Exit Function
        Case 793
            vla_step_text = "Put formula ""=A1*10"" into cell B1. [line 1346]" ' vla:4562
            Exit Function
        Case 794
            vla_step_text = "Turn off automatic calculation. [line 1347]" ' vla:4563
            Exit Function
        Case 795
            vla_step_text = "Put 3 into cell A1. [line 1348]" ' vla:4564
            Exit Function
        Case 796
            vla_step_text = "Set gc-stale to cell B1. [line 1349]" ' vla:4565
            Exit Function
        Case 797
            vla_step_text = "Put gc-stale into cell C1. [line 1350]" ' vla:4566
            Exit Function
        Case 798
            vla_step_text = "Recalculate this sheet. [line 1351]" ' vla:4567
            Exit Function
        Case 799
            vla_step_text = "Set gc-sheet to cell B1. [line 1352]" ' vla:4568
            Exit Function
        Case 800
            vla_step_text = "Put gc-sheet into cell C2. [line 1353]" ' vla:4569
            Exit Function
        Case 801
            vla_step_text = "Put 4 into cell A1. [line 1354]" ' vla:4570
            Exit Function
        Case 802
            vla_step_text = "Recalculate all open workbooks. [line 1355]" ' vla:4571
            Exit Function
        Case 803
            vla_step_text = "Set gc-all to cell B1. [line 1356]" ' vla:4572
            Exit Function
        Case 804
            vla_step_text = "Put gc-all into cell C3. [line 1357]" ' vla:4573
            Exit Function
        Case 805
            vla_step_text = "Put 5 into cell A1. [line 1358]" ' vla:4574
            Exit Function
        Case 806
            vla_step_text = "Turn on automatic calculation. [line 1359]" ' vla:4575
            Exit Function
        Case 807
            vla_step_text = "Set gc-on to cell B1. [line 1360]" ' vla:4576
            Exit Function
        Case 808
            vla_step_text = "Put gc-on into cell C4. [line 1361]" ' vla:4577
            Exit Function
        Case 809
            vla_step_text = "Put 6 into cell A1. [line 1362]" ' vla:4578
            Exit Function
        Case 810
            vla_step_text = "Set gc-auto to cell B1. [line 1363]" ' vla:4579
            Exit Function
        Case 811
            vla_step_text = "Put gc-auto into cell C5. [line 1364]" ' vla:4580
            Exit Function
        Case 812
            vla_step_text = "Check-calculation. [line 1366]" ' vla:4581
            Exit Function
        Case 813
            vla_step_text = "Work on sheet GFigures. [line 1380]" ' vla:4582
            Exit Function
        Case 814
            vla_step_text = "Clear everything from A1:J10. [line 1381]" ' vla:4583
            Exit Function
        Case 815
            vla_step_text = "Put 2 into cell B2. [line 1382]" ' vla:4584
            Exit Function
        Case 816
            vla_step_text = "Put 4 into cell B3. [line 1383]" ' vla:4585
            Exit Function
        Case 817
            vla_step_text = "Put 6 into cell B4. [line 1384]" ' vla:4586
            Exit Function
        Case 818
            vla_step_text = "Put 6 into cell B5. [line 1385]" ' vla:4587
            Exit Function
        Case 819
            vla_step_text = "Put 7 into cell B6. [line 1386]" ' vla:4588
            Exit Function
        Case 820
            vla_step_text = "Remember range B2:B6 as gfig-values. [line 1387]" ' vla:4589
            Exit Function
        Case 821
            vla_step_text = "Set gfig-middle to median of gfig-values. [line 1388]" ' vla:4590
            Exit Function
        Case 822
            vla_step_text = "Put gfig-middle into cell H1. [line 1389]" ' vla:4591
            Exit Function
        Case 823
            vla_step_text = "Put median of gfig-values into cell H2. [line 1390]" ' vla:4592
            Exit Function
        Case 824
            vla_step_text = "Set gfig-spread to standard deviation of gfig-values as a sample. [line 1391]" ' vla:4593
            Exit Function
        Case 825
            vla_step_text = "Put gfig-spread into cell H3. [line 1392]" ' vla:4594
            Exit Function
        Case 826
            vla_step_text = "Put standard deviation of gfig-values as the population into cell H4. [line 1393]" ' vla:4595
            Exit Function
        Case 827
            vla_step_text = "Put ""West"" into cell D2. [line 1394]" ' vla:4596
            Exit Function
        Case 828
            vla_step_text = "Put ""East"" into cell D3. [line 1395]" ' vla:4597
            Exit Function
        Case 829
            vla_step_text = "Put ""West"" into cell D4. [line 1396]" ' vla:4598
            Exit Function
        Case 830
            vla_step_text = "Put ""West"" into cell D5. [line 1397]" ' vla:4599
            Exit Function
        Case 831
            vla_step_text = "Put 10 into cell E2. [line 1398]" ' vla:4600
            Exit Function
        Case 832
            vla_step_text = "Put 20 into cell E3. [line 1399]" ' vla:4601
            Exit Function
        Case 833
            vla_step_text = "Put 30 into cell E5. [line 1400]" ' vla:4602
            Exit Function
        Case 834
            vla_step_text = "Put sum of range E2:E5 where range D2:D5 matches ""West"" into cell H5. [line 1401]" ' vla:4603
            Exit Function
        Case 835
            vla_step_text = "Put average of range E2:E5 where range D2:D5 matches ""West"" in cell H6. [line 1402]" ' vla:4604
            Exit Function
        Case 836
            vla_step_text = "Format cell C2 as dollars. [line 1403]" ' vla:4605
            Exit Function
        Case 837
            vla_step_text = "Put formula ""=B2*10"" into cell C2 and fill down to the last filled row of column B. [line 1404]" ' vla:4606
            Exit Function
        Case 838
            vla_step_text = "Put ""head"" into cell F1. [line 1405]" ' vla:4607
            Exit Function
        Case 839
            vla_step_text = "Put 1 into cell G2. [line 1406]" ' vla:4608
            Exit Function
        Case 840
            vla_step_text = "Put formula ""=B2*10"" into cell F2 and fill down to the last filled row of column G. [line 1407]" ' vla:4609
            Exit Function
        Case 841
            vla_step_text = "Put formula ""=B2*10"" in cell J2 and fill down to the last filled row of column I. [line 1408]" ' vla:4610
            Exit Function
        Case 842
            vla_step_text = "Check-figures. [line 1410]" ' vla:4611
            Exit Function
        Case 843
            vla_step_text = "Go to sheet Output. [line 1412]" ' vla:4612
            Exit Function
        Case 844
            vla_step_text = "Tidy-up. [line 1414]" ' vla:4613
            Exit Function
        Case 845
            vla_step_text = "Turn on screen updating. [line 1415]" ' vla:4614
            Exit Function
        Case 846
            vla_step_text = "Log ""report finished"". [line 1416]" ' vla:4615
            Exit Function
        Case Else
            vla_step_text = "an unknown step" ' vla:4616
            Exit Function
    End Select
End Function


