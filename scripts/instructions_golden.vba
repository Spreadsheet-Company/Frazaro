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
    cells(row_number, "f") = value ' vla:12 src:90
    Exit Sub ' vla:13
vla_fail: ' vla:14
    Call vla_report_error ' vla:15
End Sub

Public Function tax(amount As Variant) As Variant
    On Error GoTo vla_fail ' vla:18
    vla_step = 102 ' vla:19
    If vlatraceon() Then ' vla:20
        Call vlatracestep(102, vla_step_text(102)) ' vla:20
    End If
    tax = (amount * 0.08) ' vla:22 src:187
    Exit Function
    Exit Function ' vla:23
vla_fail: ' vla:24
    Call vla_report_error ' vla:25
End Function

Public Sub tidy_up()
    On Error GoTo vla_fail ' vla:28
    vla_step = 109 ' vla:29
    If vlatraceon() Then ' vla:30
        Call vlatracestep(109, vla_step_text(109)) ' vla:30
    End If
    Call columns("a").autofit
    vla_step = 110 ' vla:33
    If vlatraceon() Then ' vla:34
        Call vlatracestep(110, vla_step_text(110)) ' vla:34
    End If
    Call columns("b").autofit
    vla_step = 111 ' vla:37
    If vlatraceon() Then ' vla:38
        Call vlatracestep(111, vla_step_text(111)) ' vla:38
    End If
    range("a1").font.color = vlacolor(hot_pink)
    Exit Sub ' vla:41
vla_fail: ' vla:42
    Call vla_report_error ' vla:43
End Sub

Public Function commission(Optional sale As Variant = 1000, Optional rate As Variant = (5 / 100)) As Variant
    On Error GoTo vla_fail ' vla:46
    vla_step = 112 ' vla:47
    If vlatraceon() Then ' vla:48
        Call vlatracestep(112, vla_step_text(112)) ' vla:48
    End If
    commission = (sale * rate) ' vla:50 src:206
    Exit Function
    Exit Function ' vla:51
vla_fail: ' vla:52
    Call vla_report_error ' vla:53
End Function

Public Function vat_rate() As Variant
    On Error GoTo vla_fail ' vla:56
    vla_step = 122 ' vla:57
    If vlatraceon() Then ' vla:58
        Call vlatracestep(122, vla_step_text(122)) ' vla:58
    End If
    vat_rate = (20 / 100) ' vla:60 src:230
    Exit Function
    Exit Function ' vla:61
vla_fail: ' vla:62
    Call vla_report_error ' vla:63
End Function

Public Sub check_text_conditions()
    On Error GoTo vla_fail ' vla:66
    vla_step = 726 ' vla:67
    If vlatraceon() Then ' vla:68
        Call vlatracestep(726, vla_step_text(726)) ' vla:68
    End If
    If (instr(1, range("aa1"), "NAN", vbtextcompare) > 0) Then ' vla:70 src:1238
        range("l17") = "contains" ' vla:70 src:1238
    End If
    vla_step = 727 ' vla:71
    If vlatraceon() Then ' vla:72
        Call vlatracestep(727, vla_step_text(727)) ' vla:72
    End If
    If (instr(1, range("aa1"), "ban", vbtextcompare) = 1) Then ' vla:74 src:1239
        range("l18") = "starts" ' vla:74 src:1239
    End If
    vla_step = 728 ' vla:75
    If vlatraceon() Then ' vla:76
        Call vlatracestep(728, vla_step_text(728)) ' vla:76
    End If
    range("l19") = "kept" ' vla:78 src:1240
    vla_step = 729 ' vla:79
    If vlatraceon() Then ' vla:80
        Call vlatracestep(729, vla_step_text(729)) ' vla:80
    End If
    If (instr(1, range("aa1"), "nan", vbtextcompare) = 0) Then ' vla:82 src:1241
        range("l19") = "wrong" ' vla:82 src:1241
    End If
    vla_step = 730 ' vla:83
    If vlatraceon() Then ' vla:84
        Call vlatracestep(730, vla_step_text(730)) ' vla:84
    End If
    If (instr(1, range("aa1"), "kiwi", vbtextcompare) = 0) Then ' vla:86 src:1242
        range("l20") = "lacks" ' vla:86 src:1242
    End If
    vla_step = 731 ' vla:87
    If vlatraceon() Then ' vla:88
        Call vlatracestep(731, vla_step_text(731)) ' vla:88
    End If
    range("l21") = "kept" ' vla:90 src:1243
    vla_step = 732 ' vla:91
    If vlatraceon() Then ' vla:92
        Call vlatracestep(732, vla_step_text(732)) ' vla:92
    End If
    If (instr(1, range("aa1"), "kiwi", vbtextcompare) > 0) Then ' vla:94 src:1244
        range("l21") = "wrong" ' vla:94 src:1244
    End If
    vla_step = 733 ' vla:95
    If vlatraceon() Then ' vla:96
        Call vlatracestep(733, vla_step_text(733)) ' vla:96
    End If
    range("l22") = "kept" ' vla:98 src:1245
    vla_step = 734 ' vla:99
    If vlatraceon() Then ' vla:100
        Call vlatracestep(734, vla_step_text(734)) ' vla:100
    End If
    If (instr(1, range("aa1"), "nan", vbtextcompare) = 1) Then ' vla:102 src:1246
        range("l22") = "wrong" ' vla:102 src:1246
    End If
    Exit Sub ' vla:103
vla_fail: ' vla:104
    Call vla_report_error ' vla:105
End Sub

Public Sub check_formulas()
    Dim gf_last As Variant ' vla:108
    Dim gf_top As Variant ' vla:109
    Dim gf_bottom As Variant ' vla:110
    Dim gf_empty As Variant ' vla:111
    Dim gf_filled As Variant ' vla:112
    On Error GoTo vla_fail ' vla:113
    vla_step = 736 ' vla:114
    If vlatraceon() Then ' vla:115
        Call vlatracestep(736, vla_step_text(736)) ' vla:115
    End If
    Call vlaensuresheet("gformula") ' vla:117 src:1259
    Call worksheets("gformula").activate
    vla_step = 737 ' vla:118
    If vlatraceon() Then ' vla:119
        Call vlatracestep(737, vla_step_text(737)) ' vla:119
    End If
    Call range("a1:h10").clear
    vla_step = 738 ' vla:122
    If vlatraceon() Then ' vla:123
        Call vlatracestep(738, vla_step_text(738)) ' vla:123
    End If
    range("b2") = 2 ' vla:125 src:1261
    vla_step = 739 ' vla:126
    If vlatraceon() Then ' vla:127
        Call vlatracestep(739, vla_step_text(739)) ' vla:127
    End If
    range("b3") = 3 ' vla:129 src:1262
    vla_step = 740 ' vla:130
    If vlatraceon() Then ' vla:131
        Call vlatracestep(740, vla_step_text(740)) ' vla:131
    End If
    range("b4") = 4 ' vla:133 src:1263
    vla_step = 741 ' vla:134
    If vlatraceon() Then ' vla:135
        Call vlatracestep(741, vla_step_text(741)) ' vla:135
    End If
    range("c2") = 10 ' vla:137 src:1264
    vla_step = 742 ' vla:138
    If vlatraceon() Then ' vla:139
        Call vlatracestep(742, vla_step_text(742)) ' vla:139
    End If
    range("c3") = 20 ' vla:141 src:1265
    vla_step = 743 ' vla:142
    If vlatraceon() Then ' vla:143
        Call vlatracestep(743, vla_step_text(743)) ' vla:143
    End If
    range("c4") = 30 ' vla:145 src:1266
    vla_step = 744 ' vla:146
    If vlatraceon() Then ' vla:147
        Call vlatracestep(744, vla_step_text(744)) ' vla:147
    End If
    range("d2:d4").Formula2 = "=B2*C2"
    vla_step = 745 ' vla:150
    If vlatraceon() Then ' vla:151
        Call vlatracestep(745, vla_step_text(745)) ' vla:151
    End If
    gf_last = cells(rows.count, "b").end(xlup).row ' vla:153 src:1268
    vla_step = 746 ' vla:154
    If vlatraceon() Then ' vla:155
        Call vlatracestep(746, vla_step_text(746)) ' vla:155
    End If
    If (gf_last >= 2) Then
        range(("e" & 2 & ":" & "e" & gf_last)).Formula2 = "=B2+C2"
    End If
    vla_step = 747 ' vla:158
    If vlatraceon() Then ' vla:159
        Call vlatracestep(747, vla_step_text(747)) ' vla:159
    End If
    range("f1") = "head" ' vla:161 src:1270
    vla_step = 748 ' vla:162
    If vlatraceon() Then ' vla:163
        Call vlatracestep(748, vla_step_text(748)) ' vla:163
    End If
    If (1 >= 2) Then
        range(("f" & 2 & ":" & "f" & 1)).Formula2 = "=B2-C2"
    End If
    vla_step = 749 ' vla:166
    If vlatraceon() Then ' vla:167
        Call vlatracestep(749, vla_step_text(749)) ' vla:167
    End If
    gf_top = application.worksheetfunction.max(range("d2:d4")) ' vla:169 src:1272
    vla_step = 750 ' vla:170
    If vlatraceon() Then ' vla:171
        Call vlatracestep(750, vla_step_text(750)) ' vla:171
    End If
    range("h1") = gf_top ' vla:173 src:1273
    vla_step = 751 ' vla:174
    If vlatraceon() Then ' vla:175
        Call vlatracestep(751, vla_step_text(751)) ' vla:175
    End If
    gf_bottom = application.worksheetfunction.min(range("d2:d4")) ' vla:177 src:1274
    vla_step = 752 ' vla:178
    If vlatraceon() Then ' vla:179
        Call vlatracestep(752, vla_step_text(752)) ' vla:179
    End If
    range("h2") = gf_bottom ' vla:181 src:1275
    vla_step = 753 ' vla:182
    If vlatraceon() Then ' vla:183
        Call vlatracestep(753, vla_step_text(753)) ' vla:183
    End If
    range("h3") = application.worksheetfunction.average(range("b2:b4")) ' vla:185 src:1276
    vla_step = 754 ' vla:186
    If vlatraceon() Then ' vla:187
        Call vlatracestep(754, vla_step_text(754)) ' vla:187
    End If
    range("h4") = application.worksheetfunction.max(range("c2:c4")) ' vla:189 src:1277
    vla_step = 755 ' vla:190
    If vlatraceon() Then ' vla:191
        Call vlatracestep(755, vla_step_text(755)) ' vla:191
    End If
    range("h5") = application.worksheetfunction.min(range("c2:c4")) ' vla:193 src:1278
    vla_step = 756 ' vla:194
    If vlatraceon() Then ' vla:195
        Call vlatracestep(756, vla_step_text(756)) ' vla:195
    End If
    range("b6").Formula2 = "="""""
    vla_step = 757 ' vla:198
    If vlatraceon() Then ' vla:199
        Call vlatracestep(757, vla_step_text(757)) ' vla:199
    End If
    gf_empty = application.worksheetfunction.countif(range("b2:b7"), "") ' vla:201 src:1280
    vla_step = 758 ' vla:202
    If vlatraceon() Then ' vla:203
        Call vlatracestep(758, vla_step_text(758)) ' vla:203
    End If
    range("h6") = gf_empty ' vla:205 src:1281
    vla_step = 759 ' vla:206
    If vlatraceon() Then ' vla:207
        Call vlatracestep(759, vla_step_text(759)) ' vla:207
    End If
    gf_filled = application.worksheetfunction.countif(range("b2:b7"), "<>") ' vla:209 src:1282
    vla_step = 760 ' vla:210
    If vlatraceon() Then ' vla:211
        Call vlatracestep(760, vla_step_text(760)) ' vla:211
    End If
    range("h7") = gf_filled ' vla:213 src:1283
    Exit Sub ' vla:214
vla_fail: ' vla:215
    Call vla_report_error ' vla:216
End Sub

Public Sub check_statistics()
    Dim gs_middle As Variant ' vla:219
    Dim gs_spread As Variant ' vla:220
    Dim gs_west As Variant ' vla:221
    Dim gs_none As Variant ' vla:222
    On Error GoTo vla_fail ' vla:223
    vla_step = 762 ' vla:224
    If vlatraceon() Then ' vla:225
        Call vlatracestep(762, vla_step_text(762)) ' vla:225
    End If
    Call vlaensuresheet("gstats") ' vla:227 src:1300
    Call worksheets("gstats").activate
    vla_step = 763 ' vla:228
    If vlatraceon() Then ' vla:229
        Call vlatracestep(763, vla_step_text(763)) ' vla:229
    End If
    Call range("a1:h10").clear
    vla_step = 764 ' vla:232
    If vlatraceon() Then ' vla:233
        Call vlatracestep(764, vla_step_text(764)) ' vla:233
    End If
    range("b2") = 2 ' vla:235 src:1302
    vla_step = 765 ' vla:236
    If vlatraceon() Then ' vla:237
        Call vlatracestep(765, vla_step_text(765)) ' vla:237
    End If
    range("b3") = 4 ' vla:239 src:1303
    vla_step = 766 ' vla:240
    If vlatraceon() Then ' vla:241
        Call vlatracestep(766, vla_step_text(766)) ' vla:241
    End If
    range("b4") = 6 ' vla:243 src:1304
    vla_step = 767 ' vla:244
    If vlatraceon() Then ' vla:245
        Call vlatracestep(767, vla_step_text(767)) ' vla:245
    End If
    range("b5") = 6 ' vla:247 src:1305
    vla_step = 768 ' vla:248
    If vlatraceon() Then ' vla:249
        Call vlatracestep(768, vla_step_text(768)) ' vla:249
    End If
    range("b6") = 7 ' vla:251 src:1306
    vla_step = 769 ' vla:252
    If vlatraceon() Then ' vla:253
        Call vlatracestep(769, vla_step_text(769)) ' vla:253
    End If
    gs_middle = application.worksheetfunction.median(range("b2:b6")) ' vla:255 src:1307
    vla_step = 770 ' vla:256
    If vlatraceon() Then ' vla:257
        Call vlatracestep(770, vla_step_text(770)) ' vla:257
    End If
    range("h1") = gs_middle ' vla:259 src:1308
    vla_step = 771 ' vla:260
    If vlatraceon() Then ' vla:261
        Call vlatracestep(771, vla_step_text(771)) ' vla:261
    End If
    gs_spread = application.worksheetfunction.stdev_s(range("b2:b6")) ' vla:263 src:1309
    vla_step = 772 ' vla:264
    If vlatraceon() Then ' vla:265
        Call vlatracestep(772, vla_step_text(772)) ' vla:265
    End If
    range("h2") = gs_spread ' vla:267 src:1310
    vla_step = 773 ' vla:268
    If vlatraceon() Then ' vla:269
        Call vlatracestep(773, vla_step_text(773)) ' vla:269
    End If
    range("h3") = application.worksheetfunction.median(range("b2:b5")) ' vla:271 src:1311
    vla_step = 774 ' vla:272
    If vlatraceon() Then ' vla:273
        Call vlatracestep(774, vla_step_text(774)) ' vla:273
    End If
    range("h4") = application.worksheetfunction.stdev_p(range("b2:b6")) ' vla:275 src:1312
    vla_step = 775 ' vla:276
    If vlatraceon() Then ' vla:277
        Call vlatracestep(775, vla_step_text(775)) ' vla:277
    End If
    range("d2") = "West" ' vla:279 src:1313
    vla_step = 776 ' vla:280
    If vlatraceon() Then ' vla:281
        Call vlatracestep(776, vla_step_text(776)) ' vla:281
    End If
    range("d3") = "East" ' vla:283 src:1314
    vla_step = 777 ' vla:284
    If vlatraceon() Then ' vla:285
        Call vlatracestep(777, vla_step_text(777)) ' vla:285
    End If
    range("d4") = "West" ' vla:287 src:1315
    vla_step = 778 ' vla:288
    If vlatraceon() Then ' vla:289
        Call vlatracestep(778, vla_step_text(778)) ' vla:289
    End If
    range("d5") = "West" ' vla:291 src:1316
    vla_step = 779 ' vla:292
    If vlatraceon() Then ' vla:293
        Call vlatracestep(779, vla_step_text(779)) ' vla:293
    End If
    range("e2") = 10 ' vla:295 src:1317
    vla_step = 780 ' vla:296
    If vlatraceon() Then ' vla:297
        Call vlatracestep(780, vla_step_text(780)) ' vla:297
    End If
    range("e3") = 20 ' vla:299 src:1318
    vla_step = 781 ' vla:300
    If vlatraceon() Then ' vla:301
        Call vlatracestep(781, vla_step_text(781)) ' vla:301
    End If
    range("e5") = 30 ' vla:303 src:1319
    vla_step = 782 ' vla:304
    If vlatraceon() Then ' vla:305
        Call vlatracestep(782, vla_step_text(782)) ' vla:305
    End If
    gs_west = application.worksheetfunction.averageif(range("d2:d5"), "West", range("e2:e5")) ' vla:307 src:1320
    vla_step = 783 ' vla:308
    If vlatraceon() Then ' vla:309
        Call vlatracestep(783, vla_step_text(783)) ' vla:309
    End If
    range("h5") = gs_west ' vla:311 src:1321
    vla_step = 784 ' vla:312
    If vlatraceon() Then ' vla:313
        Call vlatracestep(784, vla_step_text(784)) ' vla:313
    End If
    range("h6") = "kept" ' vla:315 src:1322
    vla_step = 785 ' vla:316
    If vlatraceon() Then ' vla:317
        Call vlatracestep(785, vla_step_text(785)) ' vla:317
    End If
    On Error GoTo vla_tryf_12 ' vla:319 src:1323
    vla_step = 786 ' vla:320 src:1323
    If vlatraceon() Then ' vla:321 src:1323
        Call vlatracestep(786, vla_step_text(786)) ' vla:321 src:1323
    End If
    gs_none = application.worksheetfunction.median(range("g1:g3")) ' vla:323 src:1324
    vla_step = 787 ' vla:324 src:1323
    If vlatraceon() Then ' vla:325 src:1323
        Call vlatracestep(787, vla_step_text(787)) ' vla:325 src:1323
    End If
    range("h6") = "no stop" ' vla:327 src:1325
    GoTo vla_tryd_12 ' vla:328 src:1323
vla_tryf_12: ' vla:329 src:1323
    vla_problem = err.description ' vla:330 src:1323
    Resume vla_tryr_12 ' vla:331 src:1323
vla_tryr_12: ' vla:332 src:1323
    On Error GoTo vla_fail ' vla:333 src:1323
    vla_step = 788 ' vla:334 src:1323
    If vlatraceon() Then ' vla:335 src:1323
        Call vlatracestep(788, vla_step_text(788)) ' vla:335 src:1323
    End If
    range("h7") = vla_problem ' vla:337 src:1328
vla_tryd_12: ' vla:338 src:1323
    On Error GoTo vla_fail ' vla:339 src:1323
    Exit Sub ' vla:340
vla_fail: ' vla:341
    Call vla_report_error ' vla:342
End Sub

Public Sub check_calculation()
    Dim gc_stale As Variant ' vla:345
    Dim gc_sheet As Variant ' vla:346
    Dim gc_all As Variant ' vla:347
    Dim gc_on As Variant ' vla:348
    Dim gc_auto As Variant ' vla:349
    On Error GoTo vla_fail ' vla:350
    vla_step = 790 ' vla:351
    If vlatraceon() Then ' vla:352
        Call vlatracestep(790, vla_step_text(790)) ' vla:352
    End If
    Call vlaensuresheet("gcalc") ' vla:354 src:1343
    Call worksheets("gcalc").activate
    vla_step = 791 ' vla:355
    If vlatraceon() Then ' vla:356
        Call vlatracestep(791, vla_step_text(791)) ' vla:356
    End If
    Call range("a1:d10").clear
    vla_step = 792 ' vla:359
    If vlatraceon() Then ' vla:360
        Call vlatracestep(792, vla_step_text(792)) ' vla:360
    End If
    range("a1") = 2 ' vla:362 src:1345
    vla_step = 793 ' vla:363
    If vlatraceon() Then ' vla:364
        Call vlatracestep(793, vla_step_text(793)) ' vla:364
    End If
    range("b1").Formula2 = "=A1*10"
    vla_step = 794 ' vla:367
    If vlatraceon() Then ' vla:368
        Call vlatracestep(794, vla_step_text(794)) ' vla:368
    End If
    application.calculation = xlcalculationmanual
    vla_step = 795 ' vla:371
    If vlatraceon() Then ' vla:372
        Call vlatracestep(795, vla_step_text(795)) ' vla:372
    End If
    range("a1") = 3 ' vla:374 src:1348
    vla_step = 796 ' vla:375
    If vlatraceon() Then ' vla:376
        Call vlatracestep(796, vla_step_text(796)) ' vla:376
    End If
    gc_stale = range("b1") ' vla:378 src:1349
    vla_step = 797 ' vla:379
    If vlatraceon() Then ' vla:380
        Call vlatracestep(797, vla_step_text(797)) ' vla:380
    End If
    range("c1") = gc_stale ' vla:382 src:1350
    vla_step = 798 ' vla:383
    If vlatraceon() Then ' vla:384
        Call vlatracestep(798, vla_step_text(798)) ' vla:384
    End If
    Call activesheet.calculate
    vla_step = 799 ' vla:387
    If vlatraceon() Then ' vla:388
        Call vlatracestep(799, vla_step_text(799)) ' vla:388
    End If
    gc_sheet = range("b1") ' vla:390 src:1352
    vla_step = 800 ' vla:391
    If vlatraceon() Then ' vla:392
        Call vlatracestep(800, vla_step_text(800)) ' vla:392
    End If
    range("c2") = gc_sheet ' vla:394 src:1353
    vla_step = 801 ' vla:395
    If vlatraceon() Then ' vla:396
        Call vlatracestep(801, vla_step_text(801)) ' vla:396
    End If
    range("a1") = 4 ' vla:398 src:1354
    vla_step = 802 ' vla:399
    If vlatraceon() Then ' vla:400
        Call vlatracestep(802, vla_step_text(802)) ' vla:400
    End If
    Call application.calculate
    vla_step = 803 ' vla:403
    If vlatraceon() Then ' vla:404
        Call vlatracestep(803, vla_step_text(803)) ' vla:404
    End If
    gc_all = range("b1") ' vla:406 src:1356
    vla_step = 804 ' vla:407
    If vlatraceon() Then ' vla:408
        Call vlatracestep(804, vla_step_text(804)) ' vla:408
    End If
    range("c3") = gc_all ' vla:410 src:1357
    vla_step = 805 ' vla:411
    If vlatraceon() Then ' vla:412
        Call vlatracestep(805, vla_step_text(805)) ' vla:412
    End If
    range("a1") = 5 ' vla:414 src:1358
    vla_step = 806 ' vla:415
    If vlatraceon() Then ' vla:416
        Call vlatracestep(806, vla_step_text(806)) ' vla:416
    End If
    application.calculation = xlcalculationautomatic
    vla_step = 807 ' vla:419
    If vlatraceon() Then ' vla:420
        Call vlatracestep(807, vla_step_text(807)) ' vla:420
    End If
    gc_on = range("b1") ' vla:422 src:1360
    vla_step = 808 ' vla:423
    If vlatraceon() Then ' vla:424
        Call vlatracestep(808, vla_step_text(808)) ' vla:424
    End If
    range("c4") = gc_on ' vla:426 src:1361
    vla_step = 809 ' vla:427
    If vlatraceon() Then ' vla:428
        Call vlatracestep(809, vla_step_text(809)) ' vla:428
    End If
    range("a1") = 6 ' vla:430 src:1362
    vla_step = 810 ' vla:431
    If vlatraceon() Then ' vla:432
        Call vlatracestep(810, vla_step_text(810)) ' vla:432
    End If
    gc_auto = range("b1") ' vla:434 src:1363
    vla_step = 811 ' vla:435
    If vlatraceon() Then ' vla:436
        Call vlatracestep(811, vla_step_text(811)) ' vla:436
    End If
    range("c5") = gc_auto ' vla:438 src:1364
    Exit Sub ' vla:439
vla_fail: ' vla:440
    Call vla_report_error ' vla:441
End Sub

Public Sub check_figures()
    Dim gfig_values As Variant ' vla:444
    Dim gfig_middle As Variant ' vla:445
    Dim gfig_spread As Variant ' vla:446
    On Error GoTo vla_fail ' vla:447
    vla_step = 813 ' vla:448
    If vlatraceon() Then ' vla:449
        Call vlatracestep(813, vla_step_text(813)) ' vla:449
    End If
    Call vlaensuresheet("gfigures") ' vla:451 src:1380
    Call worksheets("gfigures").activate
    vla_step = 814 ' vla:452
    If vlatraceon() Then ' vla:453
        Call vlatracestep(814, vla_step_text(814)) ' vla:453
    End If
    Call range("a1:j10").clear
    vla_step = 815 ' vla:456
    If vlatraceon() Then ' vla:457
        Call vlatracestep(815, vla_step_text(815)) ' vla:457
    End If
    range("b2") = 2 ' vla:459 src:1382
    vla_step = 816 ' vla:460
    If vlatraceon() Then ' vla:461
        Call vlatracestep(816, vla_step_text(816)) ' vla:461
    End If
    range("b3") = 4 ' vla:463 src:1383
    vla_step = 817 ' vla:464
    If vlatraceon() Then ' vla:465
        Call vlatracestep(817, vla_step_text(817)) ' vla:465
    End If
    range("b4") = 6 ' vla:467 src:1384
    vla_step = 818 ' vla:468
    If vlatraceon() Then ' vla:469
        Call vlatracestep(818, vla_step_text(818)) ' vla:469
    End If
    range("b5") = 6 ' vla:471 src:1385
    vla_step = 819 ' vla:472
    If vlatraceon() Then ' vla:473
        Call vlatracestep(819, vla_step_text(819)) ' vla:473
    End If
    range("b6") = 7 ' vla:475 src:1386
    vla_step = 820 ' vla:476
    If vlatraceon() Then ' vla:477
        Call vlatracestep(820, vla_step_text(820)) ' vla:477
    End If
    Set gfig_values = range("b2:b6") ' vla:479 src:1387
    vla_step = 821 ' vla:480
    If vlatraceon() Then ' vla:481
        Call vlatracestep(821, vla_step_text(821)) ' vla:481
    End If
    gfig_middle = application.worksheetfunction.median(gfig_values) ' vla:483 src:1388
    vla_step = 822 ' vla:484
    If vlatraceon() Then ' vla:485
        Call vlatracestep(822, vla_step_text(822)) ' vla:485
    End If
    range("h1") = gfig_middle ' vla:487 src:1389
    vla_step = 823 ' vla:488
    If vlatraceon() Then ' vla:489
        Call vlatracestep(823, vla_step_text(823)) ' vla:489
    End If
    range("h2") = application.worksheetfunction.median(gfig_values) ' vla:491 src:1390
    vla_step = 824 ' vla:492
    If vlatraceon() Then ' vla:493
        Call vlatracestep(824, vla_step_text(824)) ' vla:493
    End If
    gfig_spread = application.worksheetfunction.stdev_s(gfig_values) ' vla:495 src:1391
    vla_step = 825 ' vla:496
    If vlatraceon() Then ' vla:497
        Call vlatracestep(825, vla_step_text(825)) ' vla:497
    End If
    range("h3") = gfig_spread ' vla:499 src:1392
    vla_step = 826 ' vla:500
    If vlatraceon() Then ' vla:501
        Call vlatracestep(826, vla_step_text(826)) ' vla:501
    End If
    range("h4") = application.worksheetfunction.stdev_p(gfig_values) ' vla:503 src:1393
    vla_step = 827 ' vla:504
    If vlatraceon() Then ' vla:505
        Call vlatracestep(827, vla_step_text(827)) ' vla:505
    End If
    range("d2") = "West" ' vla:507 src:1394
    vla_step = 828 ' vla:508
    If vlatraceon() Then ' vla:509
        Call vlatracestep(828, vla_step_text(828)) ' vla:509
    End If
    range("d3") = "East" ' vla:511 src:1395
    vla_step = 829 ' vla:512
    If vlatraceon() Then ' vla:513
        Call vlatracestep(829, vla_step_text(829)) ' vla:513
    End If
    range("d4") = "West" ' vla:515 src:1396
    vla_step = 830 ' vla:516
    If vlatraceon() Then ' vla:517
        Call vlatracestep(830, vla_step_text(830)) ' vla:517
    End If
    range("d5") = "West" ' vla:519 src:1397
    vla_step = 831 ' vla:520
    If vlatraceon() Then ' vla:521
        Call vlatracestep(831, vla_step_text(831)) ' vla:521
    End If
    range("e2") = 10 ' vla:523 src:1398
    vla_step = 832 ' vla:524
    If vlatraceon() Then ' vla:525
        Call vlatracestep(832, vla_step_text(832)) ' vla:525
    End If
    range("e3") = 20 ' vla:527 src:1399
    vla_step = 833 ' vla:528
    If vlatraceon() Then ' vla:529
        Call vlatracestep(833, vla_step_text(833)) ' vla:529
    End If
    range("e5") = 30 ' vla:531 src:1400
    vla_step = 834 ' vla:532
    If vlatraceon() Then ' vla:533
        Call vlatracestep(834, vla_step_text(834)) ' vla:533
    End If
    range("h5") = application.worksheetfunction.sumif(range("d2:d5"), "West", range("e2:e5")) ' vla:535 src:1401
    vla_step = 835 ' vla:536
    If vlatraceon() Then ' vla:537
        Call vlatracestep(835, vla_step_text(835)) ' vla:537
    End If
    range("h6") = application.worksheetfunction.averageif(range("d2:d5"), "West", range("e2:e5")) ' vla:539 src:1402
    vla_step = 836 ' vla:540
    If vlatraceon() Then ' vla:541
        Call vlatracestep(836, vla_step_text(836)) ' vla:541
    End If
    range("c2").numberformat = vlanumberformatcode("dollars", 2)
    vla_step = 837 ' vla:544
    If vlatraceon() Then ' vla:545
        Call vlatracestep(837, vla_step_text(837)) ' vla:545
    End If
    range("c2").Formula2 = "=B2*10"
    If (cells(rows.count, "b").end(xlup).row > range("c2").row) Then
        Call range("c2").resize(rowsize:=((cells(rows.count, "b").end(xlup).row - range("c2").row) + 1)).filldown
    End If
    vla_step = 838 ' vla:548
    If vlatraceon() Then ' vla:549
        Call vlatracestep(838, vla_step_text(838)) ' vla:549
    End If
    range("f1") = "head" ' vla:551 src:1405
    vla_step = 839 ' vla:552
    If vlatraceon() Then ' vla:553
        Call vlatracestep(839, vla_step_text(839)) ' vla:553
    End If
    range("g2") = 1 ' vla:555 src:1406
    vla_step = 840 ' vla:556
    If vlatraceon() Then ' vla:557
        Call vlatracestep(840, vla_step_text(840)) ' vla:557
    End If
    range("f2").Formula2 = "=B2*10"
    If (cells(rows.count, "g").end(xlup).row > range("f2").row) Then
        Call range("f2").resize(rowsize:=((cells(rows.count, "g").end(xlup).row - range("f2").row) + 1)).filldown
    End If
    vla_step = 841 ' vla:560
    If vlatraceon() Then ' vla:561
        Call vlatracestep(841, vla_step_text(841)) ' vla:561
    End If
    range("j2").Formula2 = "=B2*10"
    If (cells(rows.count, "i").end(xlup).row > range("j2").row) Then
        Call range("j2").resize(rowsize:=((cells(rows.count, "i").end(xlup).row - range("j2").row) + 1)).filldown
    End If
    Exit Sub ' vla:564
vla_fail: ' vla:565
    Call vla_report_error ' vla:566
End Sub

Public Sub check_phrases()
    Dim lph_values As Variant ' vla:569
    Dim lph_gap As Variant ' vla:570
    On Error GoTo vla_fail ' vla:571
    vla_step = 843 ' vla:572
    If vlatraceon() Then ' vla:573
        Call vlatracestep(843, vla_step_text(843)) ' vla:573
    End If
    Call vlaensuresheet("lphrases") ' vla:575 src:1422
    Call worksheets("lphrases").activate
    vla_step = 844 ' vla:576
    If vlatraceon() Then ' vla:577
        Call vlatracestep(844, vla_step_text(844)) ' vla:577
    End If
    Call range("a1:h10").clear
    vla_step = 845 ' vla:580
    If vlatraceon() Then ' vla:581
        Call vlatracestep(845, vla_step_text(845)) ' vla:581
    End If
    range("b2") = 2 ' vla:583 src:1424
    vla_step = 846 ' vla:584
    If vlatraceon() Then ' vla:585
        Call vlatracestep(846, vla_step_text(846)) ' vla:585
    End If
    range("b3") = 4 ' vla:587 src:1425
    vla_step = 847 ' vla:588
    If vlatraceon() Then ' vla:589
        Call vlatracestep(847, vla_step_text(847)) ' vla:589
    End If
    range("b4") = 6 ' vla:591 src:1426
    vla_step = 848 ' vla:592
    If vlatraceon() Then ' vla:593
        Call vlatracestep(848, vla_step_text(848)) ' vla:593
    End If
    range("b5") = 6 ' vla:595 src:1427
    vla_step = 849 ' vla:596
    If vlatraceon() Then ' vla:597
        Call vlatracestep(849, vla_step_text(849)) ' vla:597
    End If
    range("b6") = 7 ' vla:599 src:1428
    vla_step = 850 ' vla:600
    If vlatraceon() Then ' vla:601
        Call vlatracestep(850, vla_step_text(850)) ' vla:601
    End If
    Set lph_values = range("b2:b6") ' vla:603 src:1429
    vla_step = 851 ' vla:604
    If vlatraceon() Then ' vla:605
        Call vlatracestep(851, vla_step_text(851)) ' vla:605
    End If
    If (application.worksheetfunction.median(lph_values) > 5) Then ' vla:607 src:1430
        range("h1") = "above" ' vla:607 src:1430
    End If
    vla_step = 852 ' vla:608
    If vlatraceon() Then ' vla:609
        Call vlatracestep(852, vla_step_text(852)) ' vla:609
    End If
    range("h2") = (application.worksheetfunction.median(lph_values) + 1) ' vla:611 src:1431
    vla_step = 853 ' vla:612
    If vlatraceon() Then ' vla:613
        Call vlatracestep(853, vla_step_text(853)) ' vla:613
    End If
    range("h3") = application.worksheetfunction.stdev_s(range("b2:b6")) ' vla:615 src:1432
    vla_step = 854 ' vla:616
    If vlatraceon() Then ' vla:617
        Call vlatracestep(854, vla_step_text(854)) ' vla:617
    End If
    lph_gap = (application.worksheetfunction.max(range("b2:b6")) - application.worksheetfunction.min(range("b2:b6"))) ' vla:619 src:1433
    vla_step = 855 ' vla:620
    If vlatraceon() Then ' vla:621
        Call vlatracestep(855, vla_step_text(855)) ' vla:621
    End If
    range("h4") = lph_gap ' vla:623 src:1434
    vla_step = 856 ' vla:624
    If vlatraceon() Then ' vla:625
        Call vlatracestep(856, vla_step_text(856)) ' vla:625
    End If
    range("h5") = (cells(rows.count, "b").end(xlup).row + 1) ' vla:627 src:1435
    vla_step = 857 ' vla:628
    If vlatraceon() Then ' vla:629
        Call vlatracestep(857, vla_step_text(857)) ' vla:629
    End If
    If (application.worksheetfunction.stdev_p(lph_values) < 2) Then ' vla:631 src:1436
        range("h6") = "narrow" ' vla:631 src:1436
    End If
    Exit Sub ' vla:632
vla_fail: ' vla:633
    Call vla_report_error ' vla:634
End Sub

Public Sub main()
    Dim counter As Variant ' vla:637
    Dim grand As Variant ' vla:638
    Dim results As Variant ' vla:639
    Dim r As Variant ' vla:640
    Dim biggest As Variant ' vla:641
    Dim probe As Variant ' vla:642
    Dim round_check As Variant ' vla:643
    Dim thousand_check As Variant ' vla:644
    Dim f_last As Variant ' vla:645
    Dim echo_row As Variant ' vla:646
    Dim check_row As Variant ' vla:647
    Dim stripe_row As Variant ' vla:648
    Dim back_row As Variant ' vla:649
    Dim search_row As Variant ' vla:650
    Dim f As Variant ' vla:651
    Dim list_count As Variant ' vla:652
    Dim region As Variant ' vla:653
    Dim fuel As Variant ' vla:654
    Dim fee As Variant ' vla:655
    Dim full_commission As Variant ' vla:656
    Dim default_commission As Variant ' vla:657
    Dim pick_check As Variant ' vla:658
    Dim ax_price As Variant ' vla:659
    Dim key_count As Variant ' vla:660
    Dim price_sum As Variant ' vla:661
    Dim k As Variant ' vla:662
    Dim pair As Variant ' vla:663
    Dim qty_values As Variant ' vla:664
    Dim q As Variant ' vla:665
    Dim gtext_row As Variant ' vla:666
    Dim gtext_code As Variant ' vla:667
    Dim gtext_part As Variant ' vla:668
    Dim gtext_list As Variant ' vla:669
    Dim gtext_found As Variant ' vla:670
    On Error GoTo vla_fail ' vla:671
    vla_step = 1 ' vla:672
    If vlatraceon() Then ' vla:673
        Call vlatracestep(1, vla_step_text(1)) ' vla:673
    End If
    Call vlaensuresheet("output") ' vla:675 src:18
    Call worksheets("output").activate
    vla_step = 2 ' vla:676
    If vlatraceon() Then ' vla:677
        Call vlatracestep(2, vla_step_text(2)) ' vla:677
    End If
    application.screenupdating = False
    vla_step = 3 ' vla:680
    If vlatraceon() Then ' vla:681
        Call vlatracestep(3, vla_step_text(3)) ' vla:681
    End If
    Dim total As Double ' vla:683 src:21
    vla_step = 4 ' vla:684
    If vlatraceon() Then ' vla:685
        Call vlatracestep(4, vla_step_text(4)) ' vla:685
    End If
    total = 0 ' vla:687 src:22
    vla_step = 5 ' vla:688
    If vlatraceon() Then ' vla:689
        Call vlatracestep(5, vla_step_text(5)) ' vla:689
    End If
    range("a1") = "Test Report" ' vla:691 src:23
    vla_step = 6 ' vla:692
    If vlatraceon() Then ' vla:693
        Call vlatracestep(6, vla_step_text(6)) ' vla:693
    End If
    range("a1").font.bold = True
    vla_step = 7 ' vla:696
    If vlatraceon() Then ' vla:697
        Call vlatracestep(7, vla_step_text(7)) ' vla:697
    End If
    range("a1").font.size = 14
    vla_step = 8 ' vla:700
    If vlatraceon() Then ' vla:701
        Call vlatracestep(8, vla_step_text(8)) ' vla:701
    End If
    range("d1") = date() ' vla:703 src:26
    vla_step = 9 ' vla:704
    If vlatraceon() Then ' vla:705
        Call vlatracestep(9, vla_step_text(9)) ' vla:705
    End If
    For counter = 1 To 5
        vla_step = 10 ' vla:708 src:28
        If vlatraceon() Then ' vla:709 src:28
            Call vlatracestep(10, vla_step_text(10)) ' vla:709 src:28
        End If
        total = (total + counter)
        vla_step = 11 ' vla:712 src:28
        If vlatraceon() Then ' vla:713 src:28
            Call vlatracestep(11, vla_step_text(11)) ' vla:713 src:28
        End If
        If ((counter Mod 2) = 0) Then ' vla:715 src:30
            Debug.Print ("even step " & counter) ' vla:715 src:30
        End If
    Next counter
    vla_step = 12 ' vla:716
    If vlatraceon() Then ' vla:717
        Call vlatracestep(12, vla_step_text(12)) ' vla:717
    End If
    range("b2") = total ' vla:719 src:32
    vla_step = 13 ' vla:720
    If vlatraceon() Then ' vla:721
        Call vlatracestep(13, vla_step_text(13)) ' vla:721
    End If
    range("b3").Formula2 = "=B2*2"
    vla_step = 14 ' vla:724
    If vlatraceon() Then ' vla:725
        Call vlatracestep(14, vla_step_text(14)) ' vla:725
    End If
    range("b4") = application.worksheetfunction.sum(range("B2:B3")) ' vla:727 src:37
    vla_step = 15 ' vla:728
    If vlatraceon() Then ' vla:729
        Call vlatracestep(15, vla_step_text(15)) ' vla:729
    End If
    grand = application.worksheetfunction.sum(range("b2:b3")) ' vla:731 src:38
    vla_step = 16 ' vla:732
    If vlatraceon() Then ' vla:733
        Call vlatracestep(16, vla_step_text(16)) ' vla:733
    End If
    Debug.Print ("grand is " & grand) ' vla:735 src:39
    vla_step = 17 ' vla:736
    If vlatraceon() Then ' vla:737
        Call vlatracestep(17, vla_step_text(17)) ' vla:737
    End If
    If (grand > 40) Then ' vla:739 src:43
        vla_step = 18 ' vla:741 src:43
        If vlatraceon() Then ' vla:742 src:43
            Call vlatracestep(18, vla_step_text(18)) ' vla:742 src:43
        End If
        range("c4") = "PASS" ' vla:744 src:44
        vla_step = 19 ' vla:745 src:43
        If vlatraceon() Then ' vla:746 src:43
            Call vlatracestep(19, vla_step_text(19)) ' vla:746 src:43
        End If
        For counter = 1 To 2
            vla_step = 20 ' vla:749 src:45
            If vlatraceon() Then ' vla:750 src:45
                Call vlatracestep(20, vla_step_text(20)) ' vla:750 src:45
            End If
            Debug.Print ("pass check " & counter) ' vla:752 src:46
        Next counter
        vla_step = 21 ' vla:753 src:43
        If vlatraceon() Then ' vla:754 src:43
            Call vlatracestep(21, vla_step_text(21)) ' vla:754 src:43
        End If
        range("c4").font.bold = True
    Else
        vla_step = 22 ' vla:758 src:43
        If vlatraceon() Then ' vla:759 src:43
            Call vlatracestep(22, vla_step_text(22)) ' vla:759 src:43
        End If
        range("c4") = "CHECK" ' vla:761 src:51
    End If
    vla_step = 23 ' vla:762
    If vlatraceon() Then ' vla:763
        Call vlatracestep(23, vla_step_text(23)) ' vla:763
    End If
    If (grand >= 45) Then ' vla:765 src:53
        range("c4").interior.color = vbyellow
    End If
    vla_step = 24 ' vla:766
    If vlatraceon() Then ' vla:767
        Call vlatracestep(24, vla_step_text(24)) ' vla:767
    End If
    Dim label As String ' vla:769 src:54
    vla_step = 25 ' vla:770
    If vlatraceon() Then ' vla:771
        Call vlatracestep(25, vla_step_text(25)) ' vla:771
    End If
    label = ("Total: " & total) ' vla:773 src:55
    vla_step = 26 ' vla:774
    If vlatraceon() Then ' vla:775
        Call vlatracestep(26, vla_step_text(26)) ' vla:775
    End If
    range("a6") = label ' vla:777 src:56
    vla_step = 27 ' vla:778
    If vlatraceon() Then ' vla:779
        Call vlatracestep(27, vla_step_text(27)) ' vla:779
    End If
    range("a6").font.color = vlacolor(hot_pink)
    vla_step = 28 ' vla:782
    If vlatraceon() Then ' vla:783
        Call vlatracestep(28, vla_step_text(28)) ' vla:783
    End If
    Set results = range("b2:b4") ' vla:785 src:58
    vla_step = 29 ' vla:786
    If vlatraceon() Then ' vla:787
        Call vlatracestep(29, vla_step_text(29)) ' vla:787
    End If
    For Each r In results ' vla:789 src:59
        Debug.Print r ' vla:789 src:59
    Next r
    vla_step = 30 ' vla:790
    If vlatraceon() Then ' vla:791
        Call vlatracestep(30, vla_step_text(30)) ' vla:791
    End If
    biggest = application.worksheetfunction.max(results) ' vla:793 src:60
    vla_step = 31 ' vla:794
    If vlatraceon() Then ' vla:795
        Call vlatracestep(31, vla_step_text(31)) ' vla:795
    End If
    Debug.Print ((("largest result is " & biggest) & ", label length ") & len(label)) ' vla:797 src:61
    vla_step = 32 ' vla:798
    If vlatraceon() Then ' vla:799
        Call vlatracestep(32, vla_step_text(32)) ' vla:799
    End If
    For counter = 1 To 3
        vla_step = 33 ' vla:802 src:64
        If vlatraceon() Then ' vla:803 src:64
            Call vlatracestep(33, vla_step_text(33)) ' vla:803 src:64
        End If
        cells(counter, "e") = (counter * 10) ' vla:805 src:65
    Next counter
    vla_step = 34 ' vla:806
    If vlatraceon() Then ' vla:807
        Call vlatracestep(34, vla_step_text(34)) ' vla:807
    End If
    probe = cells(2, "e") ' vla:809 src:67
    vla_step = 35 ' vla:810
    If vlatraceon() Then ' vla:811
        Call vlatracestep(35, vla_step_text(35)) ' vla:811
    End If
    Dim num_col_check As String ' vla:813 src:68
    vla_step = 36 ' vla:814
    If vlatraceon() Then ' vla:815
        Call vlatracestep(36, vla_step_text(36)) ' vla:815
    End If
    If (cells(2, 5) = 20) Then ' vla:817 src:69
        num_col_check = "yes" ' vla:817 src:69
    End If
    vla_step = 37 ' vla:818
    If vlatraceon() Then ' vla:819
        Call vlatracestep(37, vla_step_text(37)) ' vla:819
    End If
    Dim value_word_check As String ' vla:821 src:70
    vla_step = 38 ' vla:822
    If vlatraceon() Then ' vla:823
        Call vlatracestep(38, vla_step_text(38)) ' vla:823
    End If
    If (cells(3, 5) = 30) Then ' vla:825 src:71
        value_word_check = "ok" ' vla:825 src:71
    End If
    vla_step = 39 ' vla:826
    If vlatraceon() Then ' vla:827
        Call vlatracestep(39, vla_step_text(39)) ' vla:827
    End If
    range("output!h16") = "bang" ' vla:829 src:72
    vla_step = 40 ' vla:830
    If vlatraceon() Then ' vla:831
        Call vlatracestep(40, vla_step_text(40)) ' vla:831
    End If
    range("a1").horizontalalignment = xlcenter
    vla_step = 41 ' vla:834
    If vlatraceon() Then ' vla:835
        Call vlatracestep(41, vla_step_text(41)) ' vla:835
    End If
    columns("d").columnwidth = 24
    vla_step = 42 ' vla:838
    If vlatraceon() Then ' vla:839
        Call vlatracestep(42, vla_step_text(42)) ' vla:839
    End If
    round_check = round(3.14159, 2) ' vla:841 src:75
    vla_step = 43 ' vla:842
    If vlatraceon() Then ' vla:843
        Call vlatracestep(43, vla_step_text(43)) ' vla:843
    End If
    Debug.Print ("rounded is " & round_check) ' vla:845 src:76
    vla_step = 44 ' vla:846
    If vlatraceon() Then ' vla:847
        Call vlatracestep(44, vla_step_text(44)) ' vla:847
    End If
    thousand_check = (1000 + 500) ' vla:849 src:77
    vla_step = 45 ' vla:850
    If vlatraceon() Then ' vla:851
        Call vlatracestep(45, vla_step_text(45)) ' vla:851
    End If
    Debug.Print ("thousands read as " & thousand_check) ' vla:853 src:78
    vla_step = 46 ' vla:854
    If vlatraceon() Then ' vla:855
        Call vlatracestep(46, vla_step_text(46)) ' vla:855
    End If
    Debug.Print ("probe is " & probe) ' vla:857 src:79
    vla_step = 47 ' vla:858
    If vlatraceon() Then ' vla:859
        Call vlatracestep(47, vla_step_text(47)) ' vla:859
    End If
    Dim countdown As Double ' vla:861 src:81
    vla_step = 48 ' vla:862
    If vlatraceon() Then ' vla:863
        Call vlatracestep(48, vla_step_text(48)) ' vla:863
    End If
    countdown = 3 ' vla:865 src:82
    vla_step = 49 ' vla:866
    If vlatraceon() Then ' vla:867
        Call vlatracestep(49, vla_step_text(49)) ' vla:867
    End If
    Do While (countdown > 0) ' vla:869 src:83
        vla_step = 50 ' vla:870 src:83
        If vlatraceon() Then ' vla:871 src:83
            Call vlatracestep(50, vla_step_text(50)) ' vla:871 src:83
        End If
        Debug.Print ("countdown " & countdown) ' vla:873 src:84
        vla_step = 51 ' vla:874 src:83
        If vlatraceon() Then ' vla:875 src:83
            Call vlatracestep(51, vla_step_text(51)) ' vla:875 src:83
        End If
        countdown = (countdown - 1) ' vla:877 src:85
    Loop
    vla_step = 53 ' vla:878
    If vlatraceon() Then ' vla:879
        Call vlatracestep(53, vla_step_text(53)) ' vla:879
    End If
    Call stamp ' vla:881 src:92
    vla_step = 54 ' vla:882
    If vlatraceon() Then ' vla:883
        Call vlatracestep(54, vla_step_text(54)) ' vla:883
    End If
    Call stamp(row_number:=2, value:="beta") ' vla:885 src:93
    vla_step = 55 ' vla:886
    If vlatraceon() Then ' vla:887
        Call vlatracestep(55, vla_step_text(55)) ' vla:887
    End If
    f_last = cells(rows.count, "f").end(xlup).row ' vla:889 src:94
    vla_step = 56 ' vla:890
    If vlatraceon() Then ' vla:891
        Call vlatracestep(56, vla_step_text(56)) ' vla:891
    End If
    Debug.Print ("column F filled to row " & f_last) ' vla:893 src:95
    vla_step = 57 ' vla:894
    If vlatraceon() Then ' vla:895
        Call vlatracestep(57, vla_step_text(57)) ' vla:895
    End If
    For echo_row = 1 To f_last ' vla:897 src:96
        Debug.Print ("echo " & echo_row) ' vla:897 src:96
    Next echo_row
    vla_step = 58 ' vla:898
    If vlatraceon() Then ' vla:899
        Call vlatracestep(58, vla_step_text(58)) ' vla:899
    End If
    For check_row = 1 To f_last ' vla:901 src:97
        vla_step = 59 ' vla:902 src:97
        If vlatraceon() Then ' vla:903 src:97
            Call vlatracestep(59, vla_step_text(59)) ' vla:903 src:97
        End If
        If (instr(1, cells(check_row, "f"), "ok", vbtextcompare) > 0) Then ' vla:905 src:98
            cells(check_row, "f").font.bold = True
        End If
    Next check_row
    vla_step = 60 ' vla:906
    If vlatraceon() Then ' vla:907
        Call vlatracestep(60, vla_step_text(60)) ' vla:907
    End If
    For stripe_row = 1 To f_last Step 2 ' vla:909 src:103
        vla_step = 61 ' vla:910 src:103
        If vlatraceon() Then ' vla:911 src:103
            Call vlatracestep(61, vla_step_text(61)) ' vla:911 src:103
        End If
        cells(stripe_row, "f").font.bold = True
    Next stripe_row
    vla_step = 62 ' vla:914
    If vlatraceon() Then ' vla:915
        Call vlatracestep(62, vla_step_text(62)) ' vla:915
    End If
    For back_row = f_last To 1 Step (0 - 2) ' vla:917 src:106
        vla_step = 63 ' vla:918 src:106
        If vlatraceon() Then ' vla:919 src:106
            Call vlatracestep(63, vla_step_text(63)) ' vla:919 src:106
        End If
        Debug.Print ("back-row " & back_row) ' vla:921 src:107
    Next back_row
    vla_step = 64 ' vla:922
    If vlatraceon() Then ' vla:923
        Call vlatracestep(64, vla_step_text(64)) ' vla:923
    End If
    For search_row = 1 To 10 ' vla:925 src:111
        vla_step = 65 ' vla:926 src:111
        If vlatraceon() Then ' vla:927 src:111
            Call vlatracestep(65, vla_step_text(65)) ' vla:927 src:111
        End If
        If (search_row = 3) Then ' vla:929 src:112
            Exit For ' vla:929 src:112
        End If
    Next search_row
    vla_step = 66 ' vla:930
    If vlatraceon() Then ' vla:931
        Call vlatracestep(66, vla_step_text(66)) ' vla:931
    End If
    Debug.Print ("stopped at " & search_row) ' vla:933 src:114
    vla_step = 67 ' vla:934
    If vlatraceon() Then ' vla:935
        Call vlatracestep(67, vla_step_text(67)) ' vla:935
    End If
    Dim found_items As Collection ' vla:937 src:119
    Set found_items = New Collection ' vla:937 src:119
    vla_step = 68 ' vla:938
    If vlatraceon() Then ' vla:939
        Call vlatracestep(68, vla_step_text(68)) ' vla:939
    End If
    For counter = 1 To 4
        vla_step = 69 ' vla:942 src:120
        If vlatraceon() Then ' vla:943 src:120
            Call vlatracestep(69, vla_step_text(69)) ' vla:943 src:120
        End If
        If (counter > 2) Then ' vla:945 src:121
            Call found_items.add((counter * 100)) ' vla:945 src:121
        End If
    Next counter
    vla_step = 70 ' vla:946
    If vlatraceon() Then ' vla:947
        Call vlatracestep(70, vla_step_text(70)) ' vla:947
    End If
    For Each f In found_items ' vla:949 src:123
        Debug.Print ("found " & f) ' vla:949 src:123
    Next f
    vla_step = 71 ' vla:950
    If vlatraceon() Then ' vla:951
        Call vlatracestep(71, vla_step_text(71)) ' vla:951
    End If
    list_count = vlacount(found_items) ' vla:953 src:124
    vla_step = 72 ' vla:954
    If vlatraceon() Then ' vla:955
        Call vlatracestep(72, vla_step_text(72)) ' vla:955
    End If
    Debug.Print ("list holds " & list_count) ' vla:957 src:125
    vla_step = 73 ' vla:958
    If vlatraceon() Then ' vla:959
        Call vlatracestep(73, vla_step_text(73)) ' vla:959
    End If
    Dim verdict As String ' vla:961 src:131
    vla_step = 74 ' vla:962
    If vlatraceon() Then ' vla:963
        Call vlatracestep(74, vla_step_text(74)) ' vla:963
    End If
    If (grand > 100) Then ' vla:965 src:132
        vla_step = 75 ' vla:967 src:132
        If vlatraceon() Then ' vla:968 src:132
            Call vlatracestep(75, vla_step_text(75)) ' vla:968 src:132
        End If
        verdict = "huge" ' vla:970 src:133
    ElseIf (grand > 40) Then
        vla_step = 76 ' vla:972 src:132
        If vlatraceon() Then ' vla:973 src:132
            Call vlatracestep(76, vla_step_text(76)) ' vla:973 src:132
        End If
        verdict = "solid" ' vla:975 src:136
    Else
        vla_step = 77 ' vla:977 src:132
        If vlatraceon() Then ' vla:978 src:132
            Call vlatracestep(77, vla_step_text(77)) ' vla:978 src:132
        End If
        verdict = "small" ' vla:980 src:139
    End If
    vla_step = 78 ' vla:981
    If vlatraceon() Then ' vla:982
        Call vlatracestep(78, vla_step_text(78)) ' vla:982
    End If
    Dim region_label As String ' vla:984 src:141
    vla_step = 79 ' vla:985
    If vlatraceon() Then ' vla:986
        Call vlatracestep(79, vla_step_text(79)) ' vla:986
    End If
    region = "South" ' vla:988 src:142
    vla_step = 80 ' vla:989
    If vlatraceon() Then ' vla:990
        Call vlatracestep(80, vla_step_text(80)) ' vla:990
    End If
    Select Case region ' vla:992 src:143
        Case "North"
            vla_step = 81 ' vla:994 src:143
            If vlatraceon() Then ' vla:995 src:143
                Call vlatracestep(81, vla_step_text(81)) ' vla:995 src:143
            End If
            region_label = "cold" ' vla:997 src:144
        Case "South", "East"
            vla_step = 82 ' vla:999 src:143
            If vlatraceon() Then ' vla:1000 src:143
                Call vlatracestep(82, vla_step_text(82)) ' vla:1000 src:143
            End If
            region_label = "warm" ' vla:1002 src:147
        Case Else
            vla_step = 83 ' vla:1004 src:143
            If vlatraceon() Then ' vla:1005 src:143
                Call vlatracestep(83, vla_step_text(83)) ' vla:1005 src:143
            End If
            region_label = "unknown" ' vla:1007 src:150
    End Select
    vla_step = 84 ' vla:1008
    If vlatraceon() Then ' vla:1009
        Call vlatracestep(84, vla_step_text(84)) ' vla:1009
    End If
    Debug.Print ((("verdict " & verdict) & ", region ") & region_label) ' vla:1011 src:152
    vla_step = 85 ' vla:1012
    If vlatraceon() Then ' vla:1013
        Call vlatracestep(85, vla_step_text(85)) ' vla:1013
    End If
    Dim until_count As Double ' vla:1015 src:154
    vla_step = 86 ' vla:1016
    If vlatraceon() Then ' vla:1017
        Call vlatracestep(86, vla_step_text(86)) ' vla:1017
    End If
    fuel = 3 ' vla:1019 src:155
    vla_step = 87 ' vla:1020
    If vlatraceon() Then ' vla:1021
        Call vlatracestep(87, vla_step_text(87)) ' vla:1021
    End If
    Do Until (fuel = 0) ' vla:1023 src:156
        vla_step = 88 ' vla:1024 src:156
        If vlatraceon() Then ' vla:1025 src:156
            Call vlatracestep(88, vla_step_text(88)) ' vla:1025 src:156
        End If
        fuel = (fuel - 1) ' vla:1027 src:157
        vla_step = 89 ' vla:1028 src:156
        If vlatraceon() Then ' vla:1029 src:156
            Call vlatracestep(89, vla_step_text(89)) ' vla:1029 src:156
        End If
        until_count = (until_count + 1)
    Loop
    vla_step = 90 ' vla:1032
    If vlatraceon() Then ' vla:1033
        Call vlatracestep(90, vla_step_text(90)) ' vla:1033
    End If
    Debug.Print (("repeat-until ran " & until_count) & " times") ' vla:1035 src:160
    vla_step = 91 ' vla:1036
    If vlatraceon() Then ' vla:1037
        Call vlatracestep(91, vla_step_text(91)) ' vla:1037
    End If
    Dim rescue As String ' vla:1039 src:166
    vla_step = 92 ' vla:1040
    If vlatraceon() Then ' vla:1041
        Call vlatracestep(92, vla_step_text(92)) ' vla:1041
    End If
    On Error GoTo vla_tryf_1 ' vla:1043 src:167
    vla_step = 93 ' vla:1044 src:167
    If vlatraceon() Then ' vla:1045 src:167
        Call vlatracestep(93, vla_step_text(93)) ' vla:1045 src:167
    End If
    Call worksheets("nowhere-land").activate
    vla_step = 94 ' vla:1048 src:167
    If vlatraceon() Then ' vla:1049 src:167
        Call vlatracestep(94, vla_step_text(94)) ' vla:1049 src:167
    End If
    rescue = "unreachable" ' vla:1051 src:169
    GoTo vla_tryd_1 ' vla:1052 src:167
vla_tryf_1: ' vla:1053 src:167
    vla_problem = err.description ' vla:1054 src:167
    Resume vla_tryr_1 ' vla:1055 src:167
vla_tryr_1: ' vla:1056 src:167
    On Error GoTo vla_fail ' vla:1057 src:167
    vla_step = 95 ' vla:1058 src:167
    If vlatraceon() Then ' vla:1059 src:167
        Call vlatracestep(95, vla_step_text(95)) ' vla:1059 src:167
    End If
    Debug.Print ("the problem was " & vla_problem) ' vla:1061 src:172
    vla_step = 96 ' vla:1062 src:167
    If vlatraceon() Then ' vla:1063 src:167
        Call vlatracestep(96, vla_step_text(96)) ' vla:1063 src:167
    End If
    If (Not (len(trim((vla_problem & ""))) = 0)) Then ' vla:1065 src:173
        rescue = "rescued" ' vla:1065 src:173
    End If
vla_tryd_1: ' vla:1066 src:167
    On Error GoTo vla_fail ' vla:1067 src:167
    vla_step = 97 ' vla:1068
    If vlatraceon() Then ' vla:1069
        Call vlatracestep(97, vla_step_text(97)) ' vla:1069
    End If
    Dim risk_free As Double ' vla:1071 src:175
    vla_step = 98 ' vla:1072
    If vlatraceon() Then ' vla:1073
        Call vlatracestep(98, vla_step_text(98)) ' vla:1073
    End If
    On Error GoTo vla_tryf_2 ' vla:1075 src:176
    vla_step = 99 ' vla:1076 src:176
    If vlatraceon() Then ' vla:1077 src:176
        Call vlatracestep(99, vla_step_text(99)) ' vla:1077 src:176
    End If
    risk_free = 7 ' vla:1079 src:177
    GoTo vla_tryd_2 ' vla:1080 src:176
vla_tryf_2: ' vla:1081 src:176
    vla_problem = err.description ' vla:1082 src:176
    Resume vla_tryr_2 ' vla:1083 src:176
vla_tryr_2: ' vla:1084 src:176
    On Error GoTo vla_fail ' vla:1085 src:176
    vla_step = 100 ' vla:1086 src:176
    If vlatraceon() Then ' vla:1087 src:176
        Call vlatracestep(100, vla_step_text(100)) ' vla:1087 src:176
    End If
    risk_free = -1 ' vla:1089 src:180
vla_tryd_2: ' vla:1090 src:176
    On Error GoTo vla_fail ' vla:1091 src:176
    vla_step = 101 ' vla:1092
    If vlatraceon() Then ' vla:1093
        Call vlatracestep(101, vla_step_text(101)) ' vla:1093
    End If
    Debug.Print ((("rescue " & rescue) & ", risk-free ") & risk_free) ' vla:1095 src:182
    vla_step = 103 ' vla:1096
    If vlatraceon() Then ' vla:1097
        Call vlatracestep(103, vla_step_text(103)) ' vla:1097
    End If
    fee = tax(100) ' vla:1099 src:189
    vla_step = 104 ' vla:1100
    If vlatraceon() Then ' vla:1101
        Call vlatracestep(104, vla_step_text(104)) ' vla:1101
    End If
    Debug.Print ("fee is " & fee) ' vla:1103 src:190
    vla_step = 105 ' vla:1104
    If vlatraceon() Then ' vla:1105
        Call vlatracestep(105, vla_step_text(105)) ' vla:1105
    End If
    Dim fee_size As String ' vla:1107 src:191
    vla_step = 106 ' vla:1108
    If vlatraceon() Then ' vla:1109
        Call vlatracestep(106, vla_step_text(106)) ' vla:1109
    End If
    If (tax(50) > 3) Then ' vla:1111 src:192
        vla_step = 107 ' vla:1113 src:192
        If vlatraceon() Then ' vla:1114 src:192
            Call vlatracestep(107, vla_step_text(107)) ' vla:1114 src:192
        End If
        fee_size = "big" ' vla:1116 src:193
    Else
        vla_step = 108 ' vla:1118 src:192
        If vlatraceon() Then ' vla:1119 src:192
            Call vlatracestep(108, vla_step_text(108)) ' vla:1119 src:192
        End If
        fee_size = "small" ' vla:1121 src:196
    End If
    vla_step = 113 ' vla:1122
    If vlatraceon() Then ' vla:1123
        Call vlatracestep(113, vla_step_text(113)) ' vla:1123
    End If
    full_commission = commission(sale:=2000, rate:=(10 / 100)) ' vla:1125 src:208
    vla_step = 114 ' vla:1126
    If vlatraceon() Then ' vla:1127
        Call vlatracestep(114, vla_step_text(114)) ' vla:1127
    End If
    default_commission = commission(sale:=600) ' vla:1129 src:209
    vla_step = 115 ' vla:1130
    If vlatraceon() Then ' vla:1131
        Call vlatracestep(115, vla_step_text(115)) ' vla:1131
    End If
    Debug.Print ((("commissions " & full_commission) & " / ") & default_commission) ' vla:1133 src:210
    vla_step = 116 ' vla:1134
    If vlatraceon() Then ' vla:1135
        Call vlatracestep(116, vla_step_text(116)) ' vla:1135
    End If
    On Error GoTo vla_tryf_3 ' vla:1137 src:219
    vla_step = 117 ' vla:1138 src:219
    If vlatraceon() Then ' vla:1139 src:219
        Call vlatracestep(117, vla_step_text(117)) ' vla:1139 src:219
    End If
    Call vlachecksheetname("Q1 Data")
    Call vlachecksheetabsent("Q1 Data")
    Call worksheets.add
    activesheet.name = "Q1 Data"
    GoTo vla_tryd_3 ' vla:1142 src:219
vla_tryf_3: ' vla:1143 src:219
    vla_problem = err.description ' vla:1144 src:219
    Resume vla_tryr_3 ' vla:1145 src:219
vla_tryr_3: ' vla:1146 src:219
    On Error GoTo vla_fail ' vla:1147 src:219
vla_tryd_3: ' vla:1148 src:219
    On Error GoTo vla_fail ' vla:1149 src:219
    vla_step = 118 ' vla:1150
    If vlatraceon() Then ' vla:1151
        Call vlatracestep(118, vla_step_text(118)) ' vla:1151
    End If
    Call worksheets("output").activate
    vla_step = 119 ' vla:1154
    If vlatraceon() Then ' vla:1155
        Call vlatracestep(119, vla_step_text(119)) ' vla:1155
    End If
    range("'Q1 Data'!A1") = "spaced" ' vla:1157 src:223
    vla_step = 120 ' vla:1158
    If vlatraceon() Then ' vla:1159
        Call vlatracestep(120, vla_step_text(120)) ' vla:1159
    End If
    Dim spaced_check As String ' vla:1161 src:224
    vla_step = 121 ' vla:1162
    If vlatraceon() Then ' vla:1163
        Call vlatracestep(121, vla_step_text(121)) ' vla:1163
    End If
    spaced_check = range("'Q1 Data'!A1") ' vla:1165 src:225
    vla_step = 123 ' vla:1166
    If vlatraceon() Then ' vla:1167
        Call vlatracestep(123, vla_step_text(123)) ' vla:1167
    End If
    Dim growth_check As Double ' vla:1169 src:237
    vla_step = 124 ' vla:1170
    If vlatraceon() Then ' vla:1171
        Call vlatracestep(124, vla_step_text(124)) ' vla:1171
    End If
    growth_check = 200 ' vla:1173 src:238
    vla_step = 125 ' vla:1174
    If vlatraceon() Then ' vla:1175
        Call vlatracestep(125, vla_step_text(125)) ' vla:1175
    End If
    growth_check = (growth_check * (1 + (10 / 100))) ' vla:1177 src:239
    vla_step = 126 ' vla:1178
    If vlatraceon() Then ' vla:1179
        Call vlatracestep(126, vla_step_text(126)) ' vla:1179
    End If
    growth_check = (growth_check * (1 + (50 / 100))) ' vla:1181 src:240
    vla_step = 127 ' vla:1182
    If vlatraceon() Then ' vla:1183
        Call vlatracestep(127, vla_step_text(127)) ' vla:1183
    End If
    growth_check = (growth_check * (1 + (100 / 100))) ' vla:1185 src:241
    vla_step = 128 ' vla:1186
    If vlatraceon() Then ' vla:1187
        Call vlatracestep(128, vla_step_text(128)) ' vla:1187
    End If
    growth_check = (growth_check * (1 - (75 / 100))) ' vla:1189 src:242
    vla_step = 129 ' vla:1190
    If vlatraceon() Then ' vla:1191
        Call vlatracestep(129, vla_step_text(129)) ' vla:1191
    End If
    pick_check = vlaitem(found_items, 2) ' vla:1193 src:243
    vla_step = 130 ' vla:1194
    If vlatraceon() Then ' vla:1195
        Call vlatracestep(130, vla_step_text(130)) ' vla:1195
    End If
    Debug.Print ((((("grew to " & growth_check) & ", picked ") & pick_check) & ", first ") & vlafirst(found_items)) ' vla:1197 src:244
    vla_step = 131 ' vla:1198
    If vlatraceon() Then ' vla:1199
        Call vlatracestep(131, vla_step_text(131)) ' vla:1199
    End If
    Dim quote_check As String ' vla:1201 src:245
    vla_step = 132 ' vla:1202
    If vlatraceon() Then ' vla:1203
        Call vlatracestep(132, vla_step_text(132)) ' vla:1203
    End If
    quote_check = "He said ""ok""" ' vla:1205 src:246
    vla_step = 133 ' vla:1206
    If vlatraceon() Then ' vla:1207
        Call vlatracestep(133, vla_step_text(133)) ' vla:1207
    End If
    Dim prices As Object ' vla:1209 src:260
    Set prices = vladictnew() ' vla:1209 src:260
    vla_step = 134 ' vla:1210
    If vlatraceon() Then ' vla:1211
        Call vlatracestep(134, vla_step_text(134)) ' vla:1211
    End If
    Call vladictset(prices, "ax-7", 100) ' vla:1213 src:261
    vla_step = 135 ' vla:1214
    If vlatraceon() Then ' vla:1215
        Call vlatracestep(135, vla_step_text(135)) ' vla:1215
    End If
    Call vladictset(prices, "bx-2", 250) ' vla:1217 src:262
    vla_step = 136 ' vla:1218
    If vlatraceon() Then ' vla:1219
        Call vlatracestep(136, vla_step_text(136)) ' vla:1219
    End If
    Call vladictset(prices, "AX-7", 120) ' vla:1221 src:263
    vla_step = 137 ' vla:1222
    If vlatraceon() Then ' vla:1223
        Call vlatracestep(137, vla_step_text(137)) ' vla:1223
    End If
    ax_price = vladictget(prices, "ax-7") ' vla:1225 src:264
    vla_step = 138 ' vla:1226
    If vlatraceon() Then ' vla:1227
        Call vlatracestep(138, vla_step_text(138)) ' vla:1227
    End If
    key_count = vlacount(vladictkeys(prices)) ' vla:1229 src:265
    vla_step = 139 ' vla:1230
    If vlatraceon() Then ' vla:1231
        Call vlatracestep(139, vla_step_text(139)) ' vla:1231
    End If
    price_sum = (vladictget(prices, "ax-7") + vladictget(prices, "bx-2")) ' vla:1233 src:266
    vla_step = 140 ' vla:1234
    If vlatraceon() Then ' vla:1235
        Call vlatracestep(140, vla_step_text(140)) ' vla:1235
    End If
    Dim key_list As String ' vla:1237 src:267
    vla_step = 141 ' vla:1238
    If vlatraceon() Then ' vla:1239
        Call vlatracestep(141, vla_step_text(141)) ' vla:1239
    End If
    For Each k In vladictkeys(prices) ' vla:1241 src:268
        vla_step = 142 ' vla:1242 src:268
        If vlatraceon() Then ' vla:1243 src:268
            Call vlatracestep(142, vla_step_text(142)) ' vla:1243 src:268
        End If
        key_list = (key_list & k) ' vla:1245 src:269
    Next k
    vla_step = 143 ' vla:1246
    If vlatraceon() Then ' vla:1247
        Call vlatracestep(143, vla_step_text(143)) ' vla:1247
    End If
    Dim price_verdict As String ' vla:1249 src:271
    vla_step = 144 ' vla:1250
    If vlatraceon() Then ' vla:1251
        Call vlatracestep(144, vla_step_text(144)) ' vla:1251
    End If
    If (vladictget(prices, "bx-2") > 200) Then ' vla:1253 src:272
        price_verdict = "steep" ' vla:1253 src:272
    End If
    vla_step = 145 ' vla:1254
    If vlatraceon() Then ' vla:1255
        Call vlatracestep(145, vla_step_text(145)) ' vla:1255
    End If
    Dim pair_trace As String ' vla:1257 src:273
    vla_step = 146 ' vla:1258
    If vlatraceon() Then ' vla:1259
        Call vlatracestep(146, vla_step_text(146)) ' vla:1259
    End If
    For Each pair In vladictpairs(prices) ' vla:1261 src:274
        vla_step = 147 ' vla:1262 src:274
        If vlatraceon() Then ' vla:1263 src:274
            Call vlatracestep(147, vla_step_text(147)) ' vla:1263 src:274
        End If
        If (vlapairvalue(pair) > 200) Then ' vla:1265 src:275
            pair_trace = (pair_trace & vlapairkey(pair)) ' vla:1265 src:275
        End If
    Next pair
    vla_step = 148 ' vla:1266
    If vlatraceon() Then ' vla:1267
        Call vlatracestep(148, vla_step_text(148)) ' vla:1267
    End If
    Debug.Print ((((("lookup: ax " & ax_price) & ", keys ") & key_list) & ", pairs ") & pair_trace) ' vla:1269 src:277
    vla_step = 149 ' vla:1270
    If vlatraceon() Then ' vla:1271
        Call vlatracestep(149, vla_step_text(149)) ' vla:1271
    End If
    range("j1") = "Item" ' vla:1273 src:284
    vla_step = 150 ' vla:1274
    If vlatraceon() Then ' vla:1275
        Call vlatracestep(150, vla_step_text(150)) ' vla:1275
    End If
    range("k1") = "Amount" ' vla:1277 src:285
    vla_step = 151 ' vla:1278
    If vlatraceon() Then ' vla:1279
        Call vlatracestep(151, vla_step_text(151)) ' vla:1279
    End If
    range("j2") = "Widget" ' vla:1281 src:286
    vla_step = 152 ' vla:1282
    If vlatraceon() Then ' vla:1283
        Call vlatracestep(152, vla_step_text(152)) ' vla:1283
    End If
    range("k2") = 10 ' vla:1285 src:287
    vla_step = 153 ' vla:1286
    If vlatraceon() Then ' vla:1287
        Call vlatracestep(153, vla_step_text(153)) ' vla:1287
    End If
    range("j3") = "Gadget" ' vla:1289 src:288
    vla_step = 154 ' vla:1290
    If vlatraceon() Then ' vla:1291
        Call vlatracestep(154, vla_step_text(154)) ' vla:1291
    End If
    range("k3") = 20 ' vla:1293 src:289
    vla_step = 155 ' vla:1294
    If vlatraceon() Then ' vla:1295
        Call vlatracestep(155, vla_step_text(155)) ' vla:1295
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j1:k3"), xllistobjecthasheaders:=xlyes).name = "salestable"
    vla_step = 156 ' vla:1298
    If vlatraceon() Then ' vla:1299
        Call vlatracestep(156, vla_step_text(156)) ' vla:1299
    End If
    activesheet.listobjects("salestable").tablestyle = "tablestylemedium9"
    vla_step = 157 ' vla:1302
    If vlatraceon() Then ' vla:1303
        Call vlatracestep(157, vla_step_text(157)) ' vla:1303
    End If
    activesheet.listobjects("salestable").showtotals = True
    vla_step = 158 ' vla:1306
    If vlatraceon() Then ' vla:1307
        Call vlatracestep(158, vla_step_text(158)) ' vla:1307
    End If
    range("j5") = "X" ' vla:1309 src:296
    vla_step = 159 ' vla:1310
    If vlatraceon() Then ' vla:1311
        Call vlatracestep(159, vla_step_text(159)) ' vla:1311
    End If
    range("j6") = "Y" ' vla:1313 src:297
    vla_step = 160 ' vla:1314
    If vlatraceon() Then ' vla:1315
        Call vlatracestep(160, vla_step_text(160)) ' vla:1315
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j5:j6"), xllistobjecthasheaders:=xlyes).name = "quiettable"
    vla_step = 161 ' vla:1318
    If vlatraceon() Then ' vla:1319
        Call vlatracestep(161, vla_step_text(161)) ' vla:1319
    End If
    activesheet.listobjects("quiettable").showtotals = True
    vla_step = 162 ' vla:1322
    If vlatraceon() Then ' vla:1323
        Call vlatracestep(162, vla_step_text(162)) ' vla:1323
    End If
    activesheet.listobjects("quiettable").showtotals = False
    vla_step = 163 ' vla:1326
    If vlatraceon() Then ' vla:1327
        Call vlatracestep(163, vla_step_text(163)) ' vla:1327
    End If
    range("j8") = "A" ' vla:1329 src:304
    vla_step = 164 ' vla:1330
    If vlatraceon() Then ' vla:1331
        Call vlatracestep(164, vla_step_text(164)) ' vla:1331
    End If
    range("j9") = "B" ' vla:1333 src:305
    vla_step = 165 ' vla:1334
    If vlatraceon() Then ' vla:1335
        Call vlatracestep(165, vla_step_text(165)) ' vla:1335
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j8:j9"), xllistobjecthasheaders:=xlyes).name = "temptable"
    vla_step = 166 ' vla:1338
    If vlatraceon() Then ' vla:1339
        Call vlatracestep(166, vla_step_text(166)) ' vla:1339
    End If
    Call activesheet.listobjects("temptable").unlist
    vla_step = 167 ' vla:1342
    If vlatraceon() Then ' vla:1343
        Call vlatracestep(167, vla_step_text(167)) ' vla:1343
    End If
    range("j11") = "Item" ' vla:1345 src:312
    vla_step = 168 ' vla:1346
    If vlatraceon() Then ' vla:1347
        Call vlatracestep(168, vla_step_text(168)) ' vla:1347
    End If
    range("k11") = "Qty" ' vla:1349 src:313
    vla_step = 169 ' vla:1350
    If vlatraceon() Then ' vla:1351
        Call vlatracestep(169, vla_step_text(169)) ' vla:1351
    End If
    range("j12") = "Bolt" ' vla:1353 src:314
    vla_step = 170 ' vla:1354
    If vlatraceon() Then ' vla:1355
        Call vlatracestep(170, vla_step_text(170)) ' vla:1355
    End If
    range("k12") = 5 ' vla:1357 src:315
    vla_step = 171 ' vla:1358
    If vlatraceon() Then ' vla:1359
        Call vlatracestep(171, vla_step_text(171)) ' vla:1359
    End If
    range("j13") = "Nut" ' vla:1361 src:316
    vla_step = 172 ' vla:1362
    If vlatraceon() Then ' vla:1363
        Call vlatracestep(172, vla_step_text(172)) ' vla:1363
    End If
    range("k13") = 8 ' vla:1365 src:317
    vla_step = 173 ' vla:1366
    If vlatraceon() Then ' vla:1367
        Call vlatracestep(173, vla_step_text(173)) ' vla:1367
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j11:k13"), xllistobjecthasheaders:=xlyes).name = "edittable"
    vla_step = 174 ' vla:1370
    If vlatraceon() Then ' vla:1371
        Call vlatracestep(174, vla_step_text(174)) ' vla:1371
    End If
    Call activesheet.listobjects("edittable").listrows.add
    vla_step = 175 ' vla:1374
    If vlatraceon() Then ' vla:1375
        Call vlatracestep(175, vla_step_text(175)) ' vla:1375
    End If
    Call activesheet.listobjects("edittable").listrows(1).delete
    vla_step = 176 ' vla:1378
    If vlatraceon() Then ' vla:1379
        Call vlatracestep(176, vla_step_text(176)) ' vla:1379
    End If
    Set qty_values = activesheet.listobjects("edittable").listcolumns("qty").databodyrange
    vla_step = 177 ' vla:1382
    If vlatraceon() Then ' vla:1383
        Call vlatracestep(177, vla_step_text(177)) ' vla:1383
    End If
    For Each q In qty_values ' vla:1385 src:322
        Debug.Print q ' vla:1385 src:322
    Next q
    vla_step = 178 ' vla:1386
    If vlatraceon() Then ' vla:1387
        Call vlatracestep(178, vla_step_text(178)) ' vla:1387
    End If
    range("n1") = "Region" ' vla:1389 src:332
    vla_step = 179 ' vla:1390
    If vlatraceon() Then ' vla:1391
        Call vlatracestep(179, vla_step_text(179)) ' vla:1391
    End If
    range("o1") = "Product" ' vla:1393 src:333
    vla_step = 180 ' vla:1394
    If vlatraceon() Then ' vla:1395
        Call vlatracestep(180, vla_step_text(180)) ' vla:1395
    End If
    range("p1") = "Segment" ' vla:1397 src:334
    vla_step = 181 ' vla:1398
    If vlatraceon() Then ' vla:1399
        Call vlatracestep(181, vla_step_text(181)) ' vla:1399
    End If
    range("q1") = "Channel" ' vla:1401 src:335
    vla_step = 182 ' vla:1402
    If vlatraceon() Then ' vla:1403
        Call vlatracestep(182, vla_step_text(182)) ' vla:1403
    End If
    range("r1") = "Units" ' vla:1405 src:336
    vla_step = 183 ' vla:1406
    If vlatraceon() Then ' vla:1407
        Call vlatracestep(183, vla_step_text(183)) ' vla:1407
    End If
    range("s1") = "Revenue" ' vla:1409 src:337
    vla_step = 184 ' vla:1410
    If vlatraceon() Then ' vla:1411
        Call vlatracestep(184, vla_step_text(184)) ' vla:1411
    End If
    range("n2") = "North" ' vla:1413 src:338
    vla_step = 185 ' vla:1414
    If vlatraceon() Then ' vla:1415
        Call vlatracestep(185, vla_step_text(185)) ' vla:1415
    End If
    range("o2") = "Widget" ' vla:1417 src:339
    vla_step = 186 ' vla:1418
    If vlatraceon() Then ' vla:1419
        Call vlatracestep(186, vla_step_text(186)) ' vla:1419
    End If
    range("p2") = "Retail" ' vla:1421 src:340
    vla_step = 187 ' vla:1422
    If vlatraceon() Then ' vla:1423
        Call vlatracestep(187, vla_step_text(187)) ' vla:1423
    End If
    range("q2") = "Online" ' vla:1425 src:341
    vla_step = 188 ' vla:1426
    If vlatraceon() Then ' vla:1427
        Call vlatracestep(188, vla_step_text(188)) ' vla:1427
    End If
    range("r2") = 10 ' vla:1429 src:342
    vla_step = 189 ' vla:1430
    If vlatraceon() Then ' vla:1431
        Call vlatracestep(189, vla_step_text(189)) ' vla:1431
    End If
    range("s2") = 500 ' vla:1433 src:343
    vla_step = 190 ' vla:1434
    If vlatraceon() Then ' vla:1435
        Call vlatracestep(190, vla_step_text(190)) ' vla:1435
    End If
    range("n3") = "North" ' vla:1437 src:344
    vla_step = 191 ' vla:1438
    If vlatraceon() Then ' vla:1439
        Call vlatracestep(191, vla_step_text(191)) ' vla:1439
    End If
    range("o3") = "Gadget" ' vla:1441 src:345
    vla_step = 192 ' vla:1442
    If vlatraceon() Then ' vla:1443
        Call vlatracestep(192, vla_step_text(192)) ' vla:1443
    End If
    range("p3") = "Wholesale" ' vla:1445 src:346
    vla_step = 193 ' vla:1446
    If vlatraceon() Then ' vla:1447
        Call vlatracestep(193, vla_step_text(193)) ' vla:1447
    End If
    range("q3") = "Store" ' vla:1449 src:347
    vla_step = 194 ' vla:1450
    If vlatraceon() Then ' vla:1451
        Call vlatracestep(194, vla_step_text(194)) ' vla:1451
    End If
    range("r3") = 5 ' vla:1453 src:348
    vla_step = 195 ' vla:1454
    If vlatraceon() Then ' vla:1455
        Call vlatracestep(195, vla_step_text(195)) ' vla:1455
    End If
    range("s3") = 200 ' vla:1457 src:349
    vla_step = 196 ' vla:1458
    If vlatraceon() Then ' vla:1459
        Call vlatracestep(196, vla_step_text(196)) ' vla:1459
    End If
    range("n4") = "South" ' vla:1461 src:350
    vla_step = 197 ' vla:1462
    If vlatraceon() Then ' vla:1463
        Call vlatracestep(197, vla_step_text(197)) ' vla:1463
    End If
    range("o4") = "Widget" ' vla:1465 src:351
    vla_step = 198 ' vla:1466
    If vlatraceon() Then ' vla:1467
        Call vlatracestep(198, vla_step_text(198)) ' vla:1467
    End If
    range("p4") = "Wholesale" ' vla:1469 src:352
    vla_step = 199 ' vla:1470
    If vlatraceon() Then ' vla:1471
        Call vlatracestep(199, vla_step_text(199)) ' vla:1471
    End If
    range("q4") = "Online" ' vla:1473 src:353
    vla_step = 200 ' vla:1474
    If vlatraceon() Then ' vla:1475
        Call vlatracestep(200, vla_step_text(200)) ' vla:1475
    End If
    range("r4") = 20 ' vla:1477 src:354
    vla_step = 201 ' vla:1478
    If vlatraceon() Then ' vla:1479
        Call vlatracestep(201, vla_step_text(201)) ' vla:1479
    End If
    range("s4") = 900 ' vla:1481 src:355
    vla_step = 202 ' vla:1482
    If vlatraceon() Then ' vla:1483
        Call vlatracestep(202, vla_step_text(202)) ' vla:1483
    End If
    range("n5") = "South" ' vla:1485 src:356
    vla_step = 203 ' vla:1486
    If vlatraceon() Then ' vla:1487
        Call vlatracestep(203, vla_step_text(203)) ' vla:1487
    End If
    range("o5") = "Gadget" ' vla:1489 src:357
    vla_step = 204 ' vla:1490
    If vlatraceon() Then ' vla:1491
        Call vlatracestep(204, vla_step_text(204)) ' vla:1491
    End If
    range("p5") = "Retail" ' vla:1493 src:358
    vla_step = 205 ' vla:1494
    If vlatraceon() Then ' vla:1495
        Call vlatracestep(205, vla_step_text(205)) ' vla:1495
    End If
    range("q5") = "Store" ' vla:1497 src:359
    vla_step = 206 ' vla:1498
    If vlatraceon() Then ' vla:1499
        Call vlatracestep(206, vla_step_text(206)) ' vla:1499
    End If
    range("r5") = 8 ' vla:1501 src:360
    vla_step = 207 ' vla:1502
    If vlatraceon() Then ' vla:1503
        Call vlatracestep(207, vla_step_text(207)) ' vla:1503
    End If
    range("s5") = 300 ' vla:1505 src:361
    vla_step = 208 ' vla:1506
    If vlatraceon() Then ' vla:1507
        Call vlatracestep(208, vla_step_text(208)) ' vla:1507
    End If
    range("n6") = "East" ' vla:1509 src:362
    vla_step = 209 ' vla:1510
    If vlatraceon() Then ' vla:1511
        Call vlatracestep(209, vla_step_text(209)) ' vla:1511
    End If
    range("o6") = "Widget" ' vla:1513 src:363
    vla_step = 210 ' vla:1514
    If vlatraceon() Then ' vla:1515
        Call vlatracestep(210, vla_step_text(210)) ' vla:1515
    End If
    range("p6") = "Retail" ' vla:1517 src:364
    vla_step = 211 ' vla:1518
    If vlatraceon() Then ' vla:1519
        Call vlatracestep(211, vla_step_text(211)) ' vla:1519
    End If
    range("q6") = "Online" ' vla:1521 src:365
    vla_step = 212 ' vla:1522
    If vlatraceon() Then ' vla:1523
        Call vlatracestep(212, vla_step_text(212)) ' vla:1523
    End If
    range("r6") = 12 ' vla:1525 src:366
    vla_step = 213 ' vla:1526
    If vlatraceon() Then ' vla:1527
        Call vlatracestep(213, vla_step_text(213)) ' vla:1527
    End If
    range("s6") = 600 ' vla:1529 src:367
    vla_step = 214 ' vla:1530
    If vlatraceon() Then ' vla:1531
        Call vlatracestep(214, vla_step_text(214)) ' vla:1531
    End If
    range("n7") = "East" ' vla:1533 src:368
    vla_step = 215 ' vla:1534
    If vlatraceon() Then ' vla:1535
        Call vlatracestep(215, vla_step_text(215)) ' vla:1535
    End If
    range("o7") = "Gadget" ' vla:1537 src:369
    vla_step = 216 ' vla:1538
    If vlatraceon() Then ' vla:1539
        Call vlatracestep(216, vla_step_text(216)) ' vla:1539
    End If
    range("p7") = "Wholesale" ' vla:1541 src:370
    vla_step = 217 ' vla:1542
    If vlatraceon() Then ' vla:1543
        Call vlatracestep(217, vla_step_text(217)) ' vla:1543
    End If
    range("q7") = "Store" ' vla:1545 src:371
    vla_step = 218 ' vla:1546
    If vlatraceon() Then ' vla:1547
        Call vlatracestep(218, vla_step_text(218)) ' vla:1547
    End If
    range("r7") = 6 ' vla:1549 src:372
    vla_step = 219 ' vla:1550
    If vlatraceon() Then ' vla:1551
        Call vlatracestep(219, vla_step_text(219)) ' vla:1551
    End If
    range("s7") = 250 ' vla:1553 src:373
    vla_step = 220 ' vla:1554
    If vlatraceon() Then ' vla:1555
        Call vlatracestep(220, vla_step_text(220)) ' vla:1555
    End If
    Call vlapivotcreate(range("n1:s7"), range("u1"), "salespivot")
    vla_step = 221 ' vla:1558
    If vlatraceon() Then ' vla:1559
        Call vlatracestep(221, vla_step_text(221)) ' vla:1559
    End If
    Call vlapivotsetorientation("salespivot", Array("region", "product"), "row")
    vla_step = 222 ' vla:1562
    If vlatraceon() Then ' vla:1563
        Call vlatracestep(222, vla_step_text(222)) ' vla:1563
    End If
    Call vlapivotsetorientation("salespivot", Array("segment"), "column")
    vla_step = 223 ' vla:1566
    If vlatraceon() Then ' vla:1567
        Call vlatracestep(223, vla_step_text(223)) ' vla:1567
    End If
    Call vlapivotsetorientation("salespivot", Array("channel"), "filter")
    vla_step = 224 ' vla:1570
    If vlatraceon() Then ' vla:1571
        Call vlatracestep(224, vla_step_text(224)) ' vla:1571
    End If
    Call vlapivotaddvalues("salespivot", Array("revenue"), "sum")
    vla_step = 225 ' vla:1574
    If vlatraceon() Then ' vla:1575
        Call vlatracestep(225, vla_step_text(225)) ' vla:1575
    End If
    Call vlapivotaddvalues("salespivot", Array("revenue", "units"), "count")
    vla_step = 226 ' vla:1578
    If vlatraceon() Then ' vla:1579
        Call vlatracestep(226, vla_step_text(226)) ' vla:1579
    End If
    Call vlapivotaddvalues("salespivot", Array("units"), "average")
    vla_step = 227 ' vla:1582
    If vlatraceon() Then ' vla:1583
        Call vlatracestep(227, vla_step_text(227)) ' vla:1583
    End If
    Call vlapivotcreate(range("n1:s7"), range("n20"), "fullpivot")
    Call vlapivotsetorientation("fullpivot", Array("region", "product", "channel"), "row")
    Call vlapivotsetorientation("fullpivot", Array("segment"), "column")
    Call vlapivotaddvalues("fullpivot", Array("revenue", "units"), "sum")
    vla_step = 228 ' vla:1586
    If vlatraceon() Then ' vla:1587
        Call vlatracestep(228, vla_step_text(228)) ' vla:1587
    End If
    Call vlapivotrefresh("salespivot")
    vla_step = 229 ' vla:1590
    If vlatraceon() Then ' vla:1591
        Call vlatracestep(229, vla_step_text(229)) ' vla:1591
    End If
    Call vlapivotrefreshall
    vla_step = 230 ' vla:1594
    If vlatraceon() Then ' vla:1595
        Call vlatracestep(230, vla_step_text(230)) ' vla:1595
    End If
    Call vlapivotsetshowdetail("salespivot", Array("region"), False)
    vla_step = 231 ' vla:1598
    If vlatraceon() Then ' vla:1599
        Call vlatracestep(231, vla_step_text(231)) ' vla:1599
    End If
    Call vlapivotsetshowdetail("fullpivot", Array("product"), False)
    vla_step = 232 ' vla:1602
    If vlatraceon() Then ' vla:1603
        Call vlatracestep(232, vla_step_text(232)) ' vla:1603
    End If
    Call vlapivotsetshowdetail("fullpivot", Array("product"), True)
    vla_step = 233 ' vla:1606
    If vlatraceon() Then ' vla:1607
        Call vlatracestep(233, vla_step_text(233)) ' vla:1607
    End If
    Call vlapivotsetrowlayout("salespivot", "tabular")
    vla_step = 234 ' vla:1610
    If vlatraceon() Then ' vla:1611
        Call vlatracestep(234, vla_step_text(234)) ' vla:1611
    End If
    Call vlapivotsetrowlayout("salespivot", "compact")
    vla_step = 235 ' vla:1614
    If vlatraceon() Then ' vla:1615
        Call vlatracestep(235, vla_step_text(235)) ' vla:1615
    End If
    Call vlapivotsetrowlayout("fullpivot", "outline")
    vla_step = 236 ' vla:1618
    If vlatraceon() Then ' vla:1619
        Call vlatracestep(236, vla_step_text(236)) ' vla:1619
    End If
    Call vlapivotsetsubtotals("salespivot", Array("region"), False)
    vla_step = 237 ' vla:1622
    If vlatraceon() Then ' vla:1623
        Call vlatracestep(237, vla_step_text(237)) ' vla:1623
    End If
    Call vlapivotsetsubtotals("fullpivot", Array("product"), False)
    vla_step = 238 ' vla:1626
    If vlatraceon() Then ' vla:1627
        Call vlatracestep(238, vla_step_text(238)) ' vla:1627
    End If
    Call vlapivotsetsubtotals("fullpivot", Array("product"), True)
    vla_step = 239 ' vla:1630
    If vlatraceon() Then ' vla:1631
        Call vlatracestep(239, vla_step_text(239)) ' vla:1631
    End If
    Call vlapivotsetblankline("salespivot", Array("region"), True)
    vla_step = 240 ' vla:1634
    If vlatraceon() Then ' vla:1635
        Call vlatracestep(240, vla_step_text(240)) ' vla:1635
    End If
    Call vlapivotsetblankline("fullpivot", Array("product"), True)
    vla_step = 241 ' vla:1638
    If vlatraceon() Then ' vla:1639
        Call vlatracestep(241, vla_step_text(241)) ' vla:1639
    End If
    Call vlapivotsetblankline("fullpivot", Array("product"), False)
    vla_step = 242 ' vla:1642
    If vlatraceon() Then ' vla:1643
        Call vlatracestep(242, vla_step_text(242)) ' vla:1643
    End If
    Call vlapivotsort("salespivot", "region", "descending", "")
    vla_step = 243 ' vla:1646
    If vlatraceon() Then ' vla:1647
        Call vlatracestep(243, vla_step_text(243)) ' vla:1647
    End If
    Call vlapivotsort("salespivot", "region", "ascending", "")
    vla_step = 244 ' vla:1650
    If vlatraceon() Then ' vla:1651
        Call vlatracestep(244, vla_step_text(244)) ' vla:1651
    End If
    Call vlapivotsort("fullpivot", "product", "descending", "revenue")
    vla_step = 245 ' vla:1654
    If vlatraceon() Then ' vla:1655
        Call vlatracestep(245, vla_step_text(245)) ' vla:1655
    End If
    Call vlapivotsort("fullpivot", "product", "ascending", "revenue")
    vla_step = 246 ' vla:1658
    If vlatraceon() Then ' vla:1659
        Call vlatracestep(246, vla_step_text(246)) ' vla:1659
    End If
    Call vlapivotcreate(range("n1:s7"), range("n200"), "renamemepivot")
    vla_step = 247 ' vla:1662
    If vlatraceon() Then ' vla:1663
        Call vlatracestep(247, vla_step_text(247)) ' vla:1663
    End If
    Call vlapivotrename("renamemepivot", "renamedpivot")
    vla_step = 248 ' vla:1666
    If vlatraceon() Then ' vla:1667
        Call vlatracestep(248, vla_step_text(248)) ' vla:1667
    End If
    Call vlapivotcreate(range("n1:s7"), range("n220"), "clearmepivot")
    vla_step = 249 ' vla:1670
    If vlatraceon() Then ' vla:1671
        Call vlatracestep(249, vla_step_text(249)) ' vla:1671
    End If
    Call vlapivotsetorientation("clearmepivot", Array("region"), "row")
    vla_step = 250 ' vla:1674
    If vlatraceon() Then ' vla:1675
        Call vlatracestep(250, vla_step_text(250)) ' vla:1675
    End If
    Call vlapivotaddvalues("clearmepivot", Array("revenue"), "sum")
    vla_step = 251 ' vla:1678
    If vlatraceon() Then ' vla:1679
        Call vlatracestep(251, vla_step_text(251)) ' vla:1679
    End If
    Call vlapivotclear("clearmepivot")
    vla_step = 252 ' vla:1682
    If vlatraceon() Then ' vla:1683
        Call vlatracestep(252, vla_step_text(252)) ' vla:1683
    End If
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
    range("n8") = "West" ' vla:1697 src:506
    vla_step = 256 ' vla:1698
    If vlatraceon() Then ' vla:1699
        Call vlatracestep(256, vla_step_text(256)) ' vla:1699
    End If
    range("o8") = "Widget" ' vla:1701 src:507
    vla_step = 257 ' vla:1702
    If vlatraceon() Then ' vla:1703
        Call vlatracestep(257, vla_step_text(257)) ' vla:1703
    End If
    range("p8") = "Retail" ' vla:1705 src:508
    vla_step = 258 ' vla:1706
    If vlatraceon() Then ' vla:1707
        Call vlatracestep(258, vla_step_text(258)) ' vla:1707
    End If
    range("q8") = "Online" ' vla:1709 src:509
    vla_step = 259 ' vla:1710
    If vlatraceon() Then ' vla:1711
        Call vlatracestep(259, vla_step_text(259)) ' vla:1711
    End If
    range("r8") = 15 ' vla:1713 src:510
    vla_step = 260 ' vla:1714
    If vlatraceon() Then ' vla:1715
        Call vlatracestep(260, vla_step_text(260)) ' vla:1715
    End If
    range("s8") = 700 ' vla:1717 src:511
    vla_step = 261 ' vla:1718
    If vlatraceon() Then ' vla:1719
        Call vlatracestep(261, vla_step_text(261)) ' vla:1719
    End If
    Call vlapivotcreate(range("n1:s7"), range("n260"), "sourcetestpivot")
    vla_step = 262 ' vla:1722
    If vlatraceon() Then ' vla:1723
        Call vlatracestep(262, vla_step_text(262)) ' vla:1723
    End If
    Call vlapivotsetorientation("sourcetestpivot", Array("region"), "row")
    vla_step = 263 ' vla:1726
    If vlatraceon() Then ' vla:1727
        Call vlatracestep(263, vla_step_text(263)) ' vla:1727
    End If
    Call vlapivotchangesource("sourcetestpivot", range("n1:s8"))
    vla_step = 264 ' vla:1730
    If vlatraceon() Then ' vla:1731
        Call vlatracestep(264, vla_step_text(264)) ' vla:1731
    End If
    Call vlapivotcreate(range("n1:s7"), range("n40"), "temppivot")
    vla_step = 265 ' vla:1734
    If vlatraceon() Then ' vla:1735
        Call vlatracestep(265, vla_step_text(265)) ' vla:1735
    End If
    Call vlapivotdelete("temppivot")
    vla_step = 266 ' vla:1738
    If vlatraceon() Then ' vla:1739
        Call vlatracestep(266, vla_step_text(266)) ' vla:1739
    End If
    range("h1") = grand ' vla:1741 src:524
    vla_step = 267 ' vla:1742
    If vlatraceon() Then ' vla:1743
        Call vlatracestep(267, vla_step_text(267)) ' vla:1743
    End If
    range("h2") = biggest ' vla:1745 src:525
    vla_step = 268 ' vla:1746
    If vlatraceon() Then ' vla:1747
        Call vlatracestep(268, vla_step_text(268)) ' vla:1747
    End If
    range("h3") = round_check ' vla:1749 src:526
    vla_step = 269 ' vla:1750
    If vlatraceon() Then ' vla:1751
        Call vlatracestep(269, vla_step_text(269)) ' vla:1751
    End If
    range("h4") = thousand_check ' vla:1753 src:527
    vla_step = 270 ' vla:1754
    If vlatraceon() Then ' vla:1755
        Call vlatracestep(270, vla_step_text(270)) ' vla:1755
    End If
    range("h5") = search_row ' vla:1757 src:528
    vla_step = 271 ' vla:1758
    If vlatraceon() Then ' vla:1759
        Call vlatracestep(271, vla_step_text(271)) ' vla:1759
    End If
    range("h6") = f_last ' vla:1761 src:529
    vla_step = 272 ' vla:1762
    If vlatraceon() Then ' vla:1763
        Call vlatracestep(272, vla_step_text(272)) ' vla:1763
    End If
    range("h7") = list_count ' vla:1765 src:530
    vla_step = 273 ' vla:1766
    If vlatraceon() Then ' vla:1767
        Call vlatracestep(273, vla_step_text(273)) ' vla:1767
    End If
    range("h8") = verdict ' vla:1769 src:531
    vla_step = 274 ' vla:1770
    If vlatraceon() Then ' vla:1771
        Call vlatracestep(274, vla_step_text(274)) ' vla:1771
    End If
    range("h9") = region_label ' vla:1773 src:532
    vla_step = 275 ' vla:1774
    If vlatraceon() Then ' vla:1775
        Call vlatracestep(275, vla_step_text(275)) ' vla:1775
    End If
    range("h10") = until_count ' vla:1777 src:533
    vla_step = 276 ' vla:1778
    If vlatraceon() Then ' vla:1779
        Call vlatracestep(276, vla_step_text(276)) ' vla:1779
    End If
    range("h11") = rescue ' vla:1781 src:534
    vla_step = 277 ' vla:1782
    If vlatraceon() Then ' vla:1783
        Call vlatracestep(277, vla_step_text(277)) ' vla:1783
    End If
    range("h12") = risk_free ' vla:1785 src:535
    vla_step = 278 ' vla:1786
    If vlatraceon() Then ' vla:1787
        Call vlatracestep(278, vla_step_text(278)) ' vla:1787
    End If
    range("h13") = fee ' vla:1789 src:536
    vla_step = 279 ' vla:1790
    If vlatraceon() Then ' vla:1791
        Call vlatracestep(279, vla_step_text(279)) ' vla:1791
    End If
    range("h14") = fee_size ' vla:1793 src:537
    vla_step = 280 ' vla:1794
    If vlatraceon() Then ' vla:1795
        Call vlatracestep(280, vla_step_text(280)) ' vla:1795
    End If
    range("h15") = num_col_check ' vla:1797 src:538
    vla_step = 281 ' vla:1798
    If vlatraceon() Then ' vla:1799
        Call vlatracestep(281, vla_step_text(281)) ' vla:1799
    End If
    range("h17") = value_word_check ' vla:1801 src:539
    vla_step = 282 ' vla:1802
    If vlatraceon() Then ' vla:1803
        Call vlatracestep(282, vla_step_text(282)) ' vla:1803
    End If
    range("h18") = spaced_check ' vla:1805 src:540
    vla_step = 283 ' vla:1806
    If vlatraceon() Then ' vla:1807
        Call vlatracestep(283, vla_step_text(283)) ' vla:1807
    End If
    range("h19") = full_commission ' vla:1809 src:541
    vla_step = 284 ' vla:1810
    If vlatraceon() Then ' vla:1811
        Call vlatracestep(284, vla_step_text(284)) ' vla:1811
    End If
    range("h20") = default_commission ' vla:1813 src:542
    vla_step = 285 ' vla:1814
    If vlatraceon() Then ' vla:1815
        Call vlatracestep(285, vla_step_text(285)) ' vla:1815
    End If
    range("h21") = growth_check ' vla:1817 src:543
    vla_step = 286 ' vla:1818
    If vlatraceon() Then ' vla:1819
        Call vlatracestep(286, vla_step_text(286)) ' vla:1819
    End If
    range("h22") = pick_check ' vla:1821 src:544
    vla_step = 287 ' vla:1822
    If vlatraceon() Then ' vla:1823
        Call vlatracestep(287, vla_step_text(287)) ' vla:1823
    End If
    range("h23") = quote_check ' vla:1825 src:545
    vla_step = 288 ' vla:1826
    If vlatraceon() Then ' vla:1827
        Call vlatracestep(288, vla_step_text(288)) ' vla:1827
    End If
    range("h24") = vat_rate() ' vla:1829 src:546
    vla_step = 289 ' vla:1830
    If vlatraceon() Then ' vla:1831
        Call vlatracestep(289, vla_step_text(289)) ' vla:1831
    End If
    range("h26") = ax_price ' vla:1833 src:547
    vla_step = 290 ' vla:1834
    If vlatraceon() Then ' vla:1835
        Call vlatracestep(290, vla_step_text(290)) ' vla:1835
    End If
    range("h27") = key_count ' vla:1837 src:548
    vla_step = 291 ' vla:1838
    If vlatraceon() Then ' vla:1839
        Call vlatracestep(291, vla_step_text(291)) ' vla:1839
    End If
    range("h28") = key_list ' vla:1841 src:549
    vla_step = 292 ' vla:1842
    If vlatraceon() Then ' vla:1843
        Call vlatracestep(292, vla_step_text(292)) ' vla:1843
    End If
    range("h29") = price_sum ' vla:1845 src:550
    vla_step = 293 ' vla:1846
    If vlatraceon() Then ' vla:1847
        Call vlatracestep(293, vla_step_text(293)) ' vla:1847
    End If
    range("h30") = price_verdict ' vla:1849 src:551
    vla_step = 294 ' vla:1850
    If vlatraceon() Then ' vla:1851
        Call vlatracestep(294, vla_step_text(294)) ' vla:1851
    End If
    range("h31") = pair_trace ' vla:1853 src:552
    vla_step = 295 ' vla:1854
    If vlatraceon() Then ' vla:1855
        Call vlatracestep(295, vla_step_text(295)) ' vla:1855
    End If
    range("h25") = "vla-row" ' vla:1857 src:559
    vla_step = 296 ' vla:1858
    If vlatraceon() Then ' vla:1859
        Call vlatracestep(296, vla_step_text(296)) ' vla:1859
    End If
    For counter = 1 To 2
        vla_step = 297 ' vla:1862 src:560
        If vlatraceon() Then ' vla:1863 src:560
            Call vlatracestep(297, vla_step_text(297)) ' vla:1863 src:560
        End If
        Debug.Print counter ' vla:1865 src:561
    Next counter
    vla_step = 298 ' vla:1866
    If vlatraceon() Then ' vla:1867
        Call vlatracestep(298, vla_step_text(298)) ' vla:1867
    End If
    Call vlaensuresheet("demo") ' vla:1869 src:576
    Call worksheets("demo").activate
    vla_step = 299 ' vla:1870
    If vlatraceon() Then ' vla:1871
        Call vlatracestep(299, vla_step_text(299)) ' vla:1871
    End If
    range("a1").font.italic = True
    vla_step = 300 ' vla:1874
    If vlatraceon() Then ' vla:1875
        Call vlatracestep(300, vla_step_text(300)) ' vla:1875
    End If
    range("a2").interior.color = vbred
    vla_step = 301 ' vla:1878
    If vlatraceon() Then ' vla:1879
        Call vlatracestep(301, vla_step_text(301)) ' vla:1879
    End If
    range("a3").interior.color = vbyellow
    vla_step = 302 ' vla:1882
    If vlatraceon() Then ' vla:1883
        Call vlatracestep(302, vla_step_text(302)) ' vla:1883
    End If
    range("a4").font.color = vlacolor(hot_pink)
    vla_step = 303 ' vla:1886
    If vlatraceon() Then ' vla:1887
        Call vlatracestep(303, vla_step_text(303)) ' vla:1887
    End If
    range("a5").interior.color = vlacolor("#FF69B4")
    vla_step = 304 ' vla:1890
    If vlatraceon() Then ' vla:1891
        Call vlatracestep(304, vla_step_text(304)) ' vla:1891
    End If
    range("a5").interior.colorindex = xlnone
    vla_step = 305 ' vla:1894
    If vlatraceon() Then ' vla:1895
        Call vlatracestep(305, vla_step_text(305)) ' vla:1895
    End If
    range("a1:e10").borders.linestyle = xlcontinuous
    vla_step = 306 ' vla:1898
    If vlatraceon() Then ' vla:1899
        Call vlatracestep(306, vla_step_text(306)) ' vla:1899
    End If
    range("b1").numberformat = "$#,##0.00"
    vla_step = 307 ' vla:1902
    If vlatraceon() Then ' vla:1903
        Call vlatracestep(307, vla_step_text(307)) ' vla:1903
    End If
    range("b2").numberformat = "0.0%"
    vla_step = 308 ' vla:1906
    If vlatraceon() Then ' vla:1907
        Call vlatracestep(308, vla_step_text(308)) ' vla:1907
    End If
    range("b3").numberformat = "mm/dd/yyyy"
    vla_step = 309 ' vla:1910
    If vlatraceon() Then ' vla:1911
        Call vlatracestep(309, vla_step_text(309)) ' vla:1911
    End If
    range("a1").horizontalalignment = xlcenter
    vla_step = 310 ' vla:1914
    If vlatraceon() Then ' vla:1915
        Call vlatracestep(310, vla_step_text(310)) ' vla:1915
    End If
    range("a1:e1").horizontalalignment = xlcenter
    vla_step = 311 ' vla:1918
    If vlatraceon() Then ' vla:1919
        Call vlatracestep(311, vla_step_text(311)) ' vla:1919
    End If
    range("a2:a5").horizontalalignment = xlleft
    vla_step = 312 ' vla:1922
    If vlatraceon() Then ' vla:1923
        Call vlatracestep(312, vla_step_text(312)) ' vla:1923
    End If
    range("a6").horizontalalignment = xlcenter
    vla_step = 313 ' vla:1926
    If vlatraceon() Then ' vla:1927
        Call vlatracestep(313, vla_step_text(313)) ' vla:1927
    End If
    range("c1:c5").wraptext = True
    vla_step = 314 ' vla:1930
    If vlatraceon() Then ' vla:1931
        Call vlatracestep(314, vla_step_text(314)) ' vla:1931
    End If
    range("c1:c5").wraptext = False
    vla_step = 315 ' vla:1934
    If vlatraceon() Then ' vla:1935
        Call vlatracestep(315, vla_step_text(315)) ' vla:1935
    End If
    Call range("d1:d3").merge
    vla_step = 316 ' vla:1938
    If vlatraceon() Then ' vla:1939
        Call vlatracestep(316, vla_step_text(316)) ' vla:1939
    End If
    Call range("d1:d3").unmerge
    vla_step = 317 ' vla:1942
    If vlatraceon() Then ' vla:1943
        Call vlatracestep(317, vla_step_text(317)) ' vla:1943
    End If
    rows(1).rowheight = 30
    vla_step = 318 ' vla:1946
    If vlatraceon() Then ' vla:1947
        Call vlatracestep(318, vla_step_text(318)) ' vla:1947
    End If
    columns("a").columnwidth = 20
    vla_step = 319 ' vla:1950
    If vlatraceon() Then ' vla:1951
        Call vlatracestep(319, vla_step_text(319)) ' vla:1951
    End If
    Call columns("b").insert
    vla_step = 320 ' vla:1954
    If vlatraceon() Then ' vla:1955
        Call vlatracestep(320, vla_step_text(320)) ' vla:1955
    End If
    columns("c").hidden = True
    vla_step = 321 ' vla:1958
    If vlatraceon() Then ' vla:1959
        Call vlatracestep(321, vla_step_text(321)) ' vla:1959
    End If
    columns("c").hidden = False
    vla_step = 322 ' vla:1962
    If vlatraceon() Then ' vla:1963
        Call vlatracestep(322, vla_step_text(322)) ' vla:1963
    End If
    Call columns("b").delete
    vla_step = 323 ' vla:1966
    If vlatraceon() Then ' vla:1967
        Call vlatracestep(323, vla_step_text(323)) ' vla:1967
    End If
    Call rows(5).insert
    vla_step = 324 ' vla:1970
    If vlatraceon() Then ' vla:1971
        Call vlatracestep(324, vla_step_text(324)) ' vla:1971
    End If
    Call rows(1).insert
    vla_step = 325 ' vla:1974
    If vlatraceon() Then ' vla:1975
        Call vlatracestep(325, vla_step_text(325)) ' vla:1975
    End If
    Call rows(2).delete
    vla_step = 326 ' vla:1978
    If vlatraceon() Then ' vla:1979
        Call vlatracestep(326, vla_step_text(326)) ' vla:1979
    End If
    Call rows(3).delete
    vla_step = 327 ' vla:1982
    If vlatraceon() Then ' vla:1983
        Call vlatracestep(327, vla_step_text(327)) ' vla:1983
    End If
    rows(10).hidden = True
    vla_step = 328 ' vla:1986
    If vlatraceon() Then ' vla:1987
        Call vlatracestep(328, vla_step_text(328)) ' vla:1987
    End If
    rows(10).hidden = False
    vla_step = 329 ' vla:1990
    If vlatraceon() Then ' vla:1991
        Call vlatracestep(329, vla_step_text(329)) ' vla:1991
    End If
    Call rows(2).select
    activewindow.freezepanes = True
    vla_step = 330 ' vla:1994
    If vlatraceon() Then ' vla:1995
        Call vlatracestep(330, vla_step_text(330)) ' vla:1995
    End If
    activewindow.freezepanes = False
    vla_step = 331 ' vla:1998
    If vlatraceon() Then ' vla:1999
        Call vlatracestep(331, vla_step_text(331)) ' vla:1999
    End If
    Call columns("a").autofit
    vla_step = 332 ' vla:2002
    If vlatraceon() Then ' vla:2003
        Call vlatracestep(332, vla_step_text(332)) ' vla:2003
    End If
    Call cells.entirecolumn.autofit
    vla_step = 333 ' vla:2006
    If vlatraceon() Then ' vla:2007
        Call vlatracestep(333, vla_step_text(333)) ' vla:2007
    End If
    range("g1") = "Region" ' vla:2009 src:620
    vla_step = 334 ' vla:2010
    If vlatraceon() Then ' vla:2011
        Call vlatracestep(334, vla_step_text(334)) ' vla:2011
    End If
    range("h1") = "Amount" ' vla:2013 src:621
    vla_step = 335 ' vla:2014
    If vlatraceon() Then ' vla:2015
        Call vlatracestep(335, vla_step_text(335)) ' vla:2015
    End If
    range("i1") = "Notes" ' vla:2017 src:622
    vla_step = 336 ' vla:2018
    If vlatraceon() Then ' vla:2019
        Call vlatracestep(336, vla_step_text(336)) ' vla:2019
    End If
    range("g2") = "West" ' vla:2021 src:623
    vla_step = 337 ' vla:2022
    If vlatraceon() Then ' vla:2023
        Call vlatracestep(337, vla_step_text(337)) ' vla:2023
    End If
    range("h2") = 100 ' vla:2025 src:624
    vla_step = 338 ' vla:2026
    If vlatraceon() Then ' vla:2027
        Call vlatracestep(338, vla_step_text(338)) ' vla:2027
    End If
    range("i2") = "ok" ' vla:2029 src:625
    vla_step = 339 ' vla:2030
    If vlatraceon() Then ' vla:2031
        Call vlatracestep(339, vla_step_text(339)) ' vla:2031
    End If
    range("g3") = "East" ' vla:2033 src:626
    vla_step = 340 ' vla:2034
    If vlatraceon() Then ' vla:2035
        Call vlatracestep(340, vla_step_text(340)) ' vla:2035
    End If
    range("h3") = 250 ' vla:2037 src:627
    vla_step = 341 ' vla:2038
    If vlatraceon() Then ' vla:2039
        Call vlatracestep(341, vla_step_text(341)) ' vla:2039
    End If
    range("i3") = "ok" ' vla:2041 src:628
    vla_step = 342 ' vla:2042
    If vlatraceon() Then ' vla:2043
        Call vlatracestep(342, vla_step_text(342)) ' vla:2043
    End If
    range("g4") = "West" ' vla:2045 src:629
    vla_step = 343 ' vla:2046
    If vlatraceon() Then ' vla:2047
        Call vlatracestep(343, vla_step_text(343)) ' vla:2047
    End If
    range("h4") = 100 ' vla:2049 src:630
    vla_step = 344 ' vla:2050
    If vlatraceon() Then ' vla:2051
        Call vlatracestep(344, vla_step_text(344)) ' vla:2051
    End If
    range("i4") = "dup" ' vla:2053 src:631
    vla_step = 345 ' vla:2054
    If vlatraceon() Then ' vla:2055
        Call vlatracestep(345, vla_step_text(345)) ' vla:2055
    End If
    Call range("g1:i4").sort(key1:=range("h1"), order1:=xldescending, header:=xlyes)
    vla_step = 346 ' vla:2058
    If vlatraceon() Then ' vla:2059
        Call vlatracestep(346, vla_step_text(346)) ' vla:2059
    End If
    Call range("g1:i4").autofilter(field:=1, criteria1:="West")
    vla_step = 347 ' vla:2062
    If vlatraceon() Then ' vla:2063
        Call vlatracestep(347, vla_step_text(347)) ' vla:2063
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 348 ' vla:2066
    If vlatraceon() Then ' vla:2067
        Call vlatracestep(348, vla_step_text(348)) ' vla:2067
    End If
    Call vlareplaceinrange(range("g1:i4"), "dup", "ok", "values")
    vla_step = 349 ' vla:2070
    If vlatraceon() Then ' vla:2071
        Call vlatracestep(349, vla_step_text(349)) ' vla:2071
    End If
    Call range("g1:i4").removeduplicates(columns:=1, header:=xlyes)
    vla_step = 350 ' vla:2074
    If vlatraceon() Then ' vla:2075
        Call vlatracestep(350, vla_step_text(350)) ' vla:2075
    End If
    Call vlareplaceinrange(columns("i"), "ok", "fine", "values")
    vla_step = 351 ' vla:2078
    If vlatraceon() Then ' vla:2079
        Call vlatracestep(351, vla_step_text(351)) ' vla:2079
    End If
    Call range("g1:i4").copy
    Call range("g1:i4").pastespecial(paste:=xlpastevalues)
    application.cutcopymode = False
    vla_step = 352 ' vla:2082
    If vlatraceon() Then ' vla:2083
        Call vlatracestep(352, vla_step_text(352)) ' vla:2083
    End If
    Call range("g1:i4").clearformats
    vla_step = 353 ' vla:2086
    If vlatraceon() Then ' vla:2087
        Call vlatracestep(353, vla_step_text(353)) ' vla:2087
    End If
    Call vlacheckrangename("demo_table")
    range("g1:i4").name = "demo_table"
    vla_step = 354 ' vla:2090
    If vlatraceon() Then ' vla:2091
        Call vlatracestep(354, vla_step_text(354)) ' vla:2091
    End If
    Call range("g1:i4").copy(destination:=range("k1:m4"))
    vla_step = 355 ' vla:2094
    If vlatraceon() Then ' vla:2095
        Call vlatracestep(355, vla_step_text(355)) ' vla:2095
    End If
    worksheets("demo").tab.color = vlacolor(hot_pink)
    vla_step = 356 ' vla:2098
    If vlatraceon() Then ' vla:2099
        Call vlatracestep(356, vla_step_text(356)) ' vla:2099
    End If
    Call activesheet.protect(password:="demo123")
    vla_step = 357 ' vla:2102
    If vlatraceon() Then ' vla:2103
        Call vlatracestep(357, vla_step_text(357)) ' vla:2103
    End If
    Call activesheet.unprotect(password:="demo123")
    vla_step = 358 ' vla:2106
    If vlatraceon() Then ' vla:2107
        Call vlatracestep(358, vla_step_text(358)) ' vla:2107
    End If
    On Error GoTo vla_tryf_4 ' vla:2109 src:675
    vla_step = 359 ' vla:2110 src:675
    If vlatraceon() Then ' vla:2111 src:675
        Call vlatracestep(359, vla_step_text(359)) ' vla:2111 src:675
    End If
    application.displayalerts = False
    Call worksheets("gstruct").delete
    application.displayalerts = True
    GoTo vla_tryd_4 ' vla:2114 src:675
vla_tryf_4: ' vla:2115 src:675
    vla_problem = err.description ' vla:2116 src:675
    Resume vla_tryr_4 ' vla:2117 src:675
vla_tryr_4: ' vla:2118 src:675
    On Error GoTo vla_fail ' vla:2119 src:675
vla_tryd_4: ' vla:2120 src:675
    On Error GoTo vla_fail ' vla:2121 src:675
    vla_step = 360 ' vla:2122
    If vlatraceon() Then ' vla:2123
        Call vlatracestep(360, vla_step_text(360)) ' vla:2123
    End If
    Call vlaensuresheet("gstruct") ' vla:2125 src:678
    Call worksheets("gstruct").activate
    vla_step = 361 ' vla:2126
    If vlatraceon() Then ' vla:2127
        Call vlatracestep(361, vla_step_text(361)) ' vla:2127
    End If
    rows(3).hidden = True
    vla_step = 362 ' vla:2130
    If vlatraceon() Then ' vla:2131
        Call vlatracestep(362, vla_step_text(362)) ' vla:2131
    End If
    columns("b").hidden = True
    vla_step = 363 ' vla:2134
    If vlatraceon() Then ' vla:2135
        Call vlatracestep(363, vla_step_text(363)) ' vla:2135
    End If
    cells.entirerow.hidden = False
    cells.entirecolumn.hidden = False
    vla_step = 364 ' vla:2138
    If vlatraceon() Then ' vla:2139
        Call vlatracestep(364, vla_step_text(364)) ' vla:2139
    End If
    range("a6").font.size = 36
    vla_step = 365 ' vla:2142
    If vlatraceon() Then ' vla:2143
        Call vlatracestep(365, vla_step_text(365)) ' vla:2143
    End If
    Call rows(6).autofit
    vla_step = 366 ' vla:2146
    If vlatraceon() Then ' vla:2147
        Call vlatracestep(366, vla_step_text(366)) ' vla:2147
    End If
    Call rows((10 & ":" & 12)).group
    vla_step = 367 ' vla:2150
    If vlatraceon() Then ' vla:2151
        Call vlatracestep(367, vla_step_text(367)) ' vla:2151
    End If
    Call rows((14 & ":" & 16)).group
    vla_step = 368 ' vla:2154
    If vlatraceon() Then ' vla:2155
        Call vlatracestep(368, vla_step_text(368)) ' vla:2155
    End If
    Call rows((14 & ":" & 16)).ungroup
    vla_step = 369 ' vla:2158
    If vlatraceon() Then ' vla:2159
        Call vlatracestep(369, vla_step_text(369)) ' vla:2159
    End If
    range("a20") = "before-insert" ' vla:2161 src:691
    vla_step = 370 ' vla:2162
    If vlatraceon() Then ' vla:2163
        Call vlatracestep(370, vla_step_text(370)) ' vla:2163
    End If
    Call rows(20).resize(rowsize:=3).insert
    vla_step = 371 ' vla:2166
    If vlatraceon() Then ' vla:2167
        Call vlatracestep(371, vla_step_text(371)) ' vla:2167
    End If
    range("a40") = "before-delete" ' vla:2169 src:694
    vla_step = 372 ' vla:2170
    If vlatraceon() Then ' vla:2171
        Call vlatracestep(372, vla_step_text(372)) ' vla:2171
    End If
    Call rows((38 & ":" & 39)).delete
    vla_step = 373 ' vla:2174
    If vlatraceon() Then ' vla:2175
        Call vlatracestep(373, vla_step_text(373)) ' vla:2175
    End If
    range("b50") = "marker-b" ' vla:2177 src:697
    vla_step = 374 ' vla:2178
    If vlatraceon() Then ' vla:2179
        Call vlatracestep(374, vla_step_text(374)) ' vla:2179
    End If
    range("c50") = "marker-c" ' vla:2181 src:698
    vla_step = 375 ' vla:2182
    If vlatraceon() Then ' vla:2183
        Call vlatracestep(375, vla_step_text(375)) ' vla:2183
    End If
    range("d50") = "marker-d" ' vla:2185 src:699
    vla_step = 376 ' vla:2186
    If vlatraceon() Then ' vla:2187
        Call vlatracestep(376, vla_step_text(376)) ' vla:2187
    End If
    Call vlamovecolumn("b", "d")
    vla_step = 377 ' vla:2190
    If vlatraceon() Then ' vla:2191
        Call vlatracestep(377, vla_step_text(377)) ' vla:2191
    End If
    Call vlafreezepanes(2)
    vla_step = 378 ' vla:2194
    If vlatraceon() Then ' vla:2195
        Call vlatracestep(378, vla_step_text(378)) ' vla:2195
    End If
    range("a60").font.bold = True
    vla_step = 379 ' vla:2198
    If vlatraceon() Then ' vla:2199
        Call vlatracestep(379, vla_step_text(379)) ' vla:2199
    End If
    range("a60").interior.color = vbred
    vla_step = 380 ' vla:2202
    If vlatraceon() Then ' vla:2203
        Call vlatracestep(380, vla_step_text(380)) ' vla:2203
    End If
    Call range("a60:a60").copy
    Call range("c60:c60").pastespecial(paste:=xlpasteformats)
    application.cutcopymode = False
    vla_step = 381 ' vla:2206
    If vlatraceon() Then ' vla:2207
        Call vlatracestep(381, vla_step_text(381)) ' vla:2207
    End If
    Call range("a60:a60").copy
    Call range("c62:c62").pastespecial(paste:=xlpasteformats)
    application.cutcopymode = False
    vla_step = 382 ' vla:2210
    If vlatraceon() Then ' vla:2211
        Call vlatracestep(382, vla_step_text(382)) ' vla:2211
    End If
    range("b64").Formula2 = "=5*2"
    vla_step = 383 ' vla:2214
    If vlatraceon() Then ' vla:2215
        Call vlatracestep(383, vla_step_text(383)) ' vla:2215
    End If
    Call range("b64").copy
    Call range("d64").pastespecial(paste:=xlpasteformulas)
    application.cutcopymode = False
    vla_step = 384 ' vla:2218
    If vlatraceon() Then ' vla:2219
        Call vlatracestep(384, vla_step_text(384)) ' vla:2219
    End If
    columns("f").columnwidth = 33
    vla_step = 385 ' vla:2222
    If vlatraceon() Then ' vla:2223
        Call vlatracestep(385, vla_step_text(385)) ' vla:2223
    End If
    Call range("f1:f1").copy
    Call range("h1:h1").pastespecial(paste:=xlpastecolumnwidths)
    application.cutcopymode = False
    vla_step = 386 ' vla:2226
    If vlatraceon() Then ' vla:2227
        Call vlatracestep(386, vla_step_text(386)) ' vla:2227
    End If
    range("a68") = 1 ' vla:2229 src:716
    vla_step = 387 ' vla:2230
    If vlatraceon() Then ' vla:2231
        Call vlatracestep(387, vla_step_text(387)) ' vla:2231
    End If
    range("a69") = 2 ' vla:2233 src:717
    vla_step = 388 ' vla:2234
    If vlatraceon() Then ' vla:2235
        Call vlatracestep(388, vla_step_text(388)) ' vla:2235
    End If
    range("a70") = 3 ' vla:2237 src:718
    vla_step = 389 ' vla:2238
    If vlatraceon() Then ' vla:2239
        Call vlatracestep(389, vla_step_text(389)) ' vla:2239
    End If
    Call range("a68:a70").copy
    Call range("c68").pastespecial(transpose:=True)
    application.cutcopymode = False
    vla_step = 390 ' vla:2242
    If vlatraceon() Then ' vla:2243
        Call vlatracestep(390, vla_step_text(390)) ' vla:2243
    End If
    range("a72") = "cutme" ' vla:2245 src:721
    vla_step = 391 ' vla:2246
    If vlatraceon() Then ' vla:2247
        Call vlatracestep(391, vla_step_text(391)) ' vla:2247
    End If
    Call range("a72:a72").cut(destination:=range("c72"))
    vla_step = 392 ' vla:2250
    If vlatraceon() Then ' vla:2251
        Call vlatracestep(392, vla_step_text(392)) ' vla:2251
    End If
    range("a74") = "rowdata" ' vla:2253 src:724
    vla_step = 393 ' vla:2254
    If vlatraceon() Then ' vla:2255
        Call vlatracestep(393, vla_step_text(393)) ' vla:2255
    End If
    Call rows(74).copy(destination:=rows(76))
    vla_step = 394 ' vla:2258
    If vlatraceon() Then ' vla:2259
        Call vlatracestep(394, vla_step_text(394)) ' vla:2259
    End If
    range("a80") = "clearme" ' vla:2261 src:728
    vla_step = 395 ' vla:2262
    If vlatraceon() Then ' vla:2263
        Call vlatracestep(395, vla_step_text(395)) ' vla:2263
    End If
    range("a80").font.bold = True
    vla_step = 396 ' vla:2266
    If vlatraceon() Then ' vla:2267
        Call vlatracestep(396, vla_step_text(396)) ' vla:2267
    End If
    range("a80").interior.color = vbred
    vla_step = 397 ' vla:2270
    If vlatraceon() Then ' vla:2271
        Call vlatracestep(397, vla_step_text(397)) ' vla:2271
    End If
    Call range("a80:b80").clear
    vla_step = 398 ' vla:2274
    If vlatraceon() Then ' vla:2275
        Call vlatracestep(398, vla_step_text(398)) ' vla:2275
    End If
    range("a84") = "m1" ' vla:2277 src:733
    vla_step = 399 ' vla:2278
    If vlatraceon() Then ' vla:2279
        Call vlatracestep(399, vla_step_text(399)) ' vla:2279
    End If
    range("a85") = "m2" ' vla:2281 src:734
    vla_step = 400 ' vla:2282
    If vlatraceon() Then ' vla:2283
        Call vlatracestep(400, vla_step_text(400)) ' vla:2283
    End If
    range("a86") = "m3" ' vla:2285 src:735
    vla_step = 401 ' vla:2286
    If vlatraceon() Then ' vla:2287
        Call vlatracestep(401, vla_step_text(401)) ' vla:2287
    End If
    range("a87") = "m4" ' vla:2289 src:736
    vla_step = 402 ' vla:2290
    If vlatraceon() Then ' vla:2291
        Call vlatracestep(402, vla_step_text(402)) ' vla:2291
    End If
    range("a88") = "m5" ' vla:2293 src:737
    vla_step = 403 ' vla:2294
    If vlatraceon() Then ' vla:2295
        Call vlatracestep(403, vla_step_text(403)) ' vla:2295
    End If
    Call range("a84:a86").delete(shift:=xlshiftup)
    vla_step = 404 ' vla:2298
    If vlatraceon() Then ' vla:2299
        Call vlatracestep(404, vla_step_text(404)) ' vla:2299
    End If
    range("b90") = "x1" ' vla:2301 src:740
    vla_step = 405 ' vla:2302
    If vlatraceon() Then ' vla:2303
        Call vlatracestep(405, vla_step_text(405)) ' vla:2303
    End If
    range("c90") = "x2" ' vla:2305 src:741
    vla_step = 406 ' vla:2306
    If vlatraceon() Then ' vla:2307
        Call vlatracestep(406, vla_step_text(406)) ' vla:2307
    End If
    range("d90") = "x3" ' vla:2309 src:742
    vla_step = 407 ' vla:2310
    If vlatraceon() Then ' vla:2311
        Call vlatracestep(407, vla_step_text(407)) ' vla:2311
    End If
    range("e90") = "rightdata" ' vla:2313 src:743
    vla_step = 408 ' vla:2314
    If vlatraceon() Then ' vla:2315
        Call vlatracestep(408, vla_step_text(408)) ' vla:2315
    End If
    Call range("b90:d90").delete(shift:=xlshifttoleft)
    vla_step = 409 ' vla:2318
    If vlatraceon() Then ' vla:2319
        Call vlatracestep(409, vla_step_text(409)) ' vla:2319
    End If
    range("a94") = 1 ' vla:2321 src:746
    vla_step = 410 ' vla:2322
    If vlatraceon() Then ' vla:2323
        Call vlatracestep(410, vla_step_text(410)) ' vla:2323
    End If
    range("b94") = 2 ' vla:2325 src:747
    vla_step = 411 ' vla:2326
    If vlatraceon() Then ' vla:2327
        Call vlatracestep(411, vla_step_text(411)) ' vla:2327
    End If
    range("a95") = 3 ' vla:2329 src:748
    vla_step = 412 ' vla:2330
    If vlatraceon() Then ' vla:2331
        Call vlatracestep(412, vla_step_text(412)) ' vla:2331
    End If
    range("a96") = 5 ' vla:2333 src:749
    vla_step = 413 ' vla:2334
    If vlatraceon() Then ' vla:2335
        Call vlatracestep(413, vla_step_text(413)) ' vla:2335
    End If
    range("b96") = 6 ' vla:2337 src:750
    vla_step = 414 ' vla:2338
    If vlatraceon() Then ' vla:2339
        Call vlatracestep(414, vla_step_text(414)) ' vla:2339
    End If
    Call vladeleteblankrows(range("a94:b96"))
    vla_step = 415 ' vla:2342
    If vlatraceon() Then ' vla:2343
        Call vlatracestep(415, vla_step_text(415)) ' vla:2343
    End If
    range("a150") = 1 ' vla:2345 src:754
    vla_step = 416 ' vla:2346
    If vlatraceon() Then ' vla:2347
        Call vlatracestep(416, vla_step_text(416)) ' vla:2347
    End If
    range("a151") = 2 ' vla:2349 src:755
    vla_step = 417 ' vla:2350
    If vlatraceon() Then ' vla:2351
        Call vlatracestep(417, vla_step_text(417)) ' vla:2351
    End If
    range("a152") = 3 ' vla:2353 src:756
    vla_step = 418 ' vla:2354
    If vlatraceon() Then ' vla:2355
        Call vlatracestep(418, vla_step_text(418)) ' vla:2355
    End If
    range("a153") = 4 ' vla:2357 src:757
    vla_step = 419 ' vla:2358
    If vlatraceon() Then ' vla:2359
        Call vlatracestep(419, vla_step_text(419)) ' vla:2359
    End If
    Call vlabandrows(range("a150:d153"), vlacolor("#D9D9D9"))
    vla_step = 420 ' vla:2362
    If vlatraceon() Then ' vla:2363
        Call vlatracestep(420, vla_step_text(420)) ' vla:2363
    End If
    rows(165).font.bold = True
    rows(165).interior.color = vlacolor("#D9D9D9")
    vla_step = 421 ' vla:2366
    If vlatraceon() Then ' vla:2367
        Call vlatracestep(421, vla_step_text(421)) ' vla:2367
    End If
    Call vlafillseries(range("a170:a179"), 1, 1, "linear")
    vla_step = 422 ' vla:2370
    If vlatraceon() Then ' vla:2371
        Call vlatracestep(422, vla_step_text(422)) ' vla:2371
    End If
    Call vlafillseries(range("a180:a184"), 1, 2, "linear")
    vla_step = 423 ' vla:2374
    If vlatraceon() Then ' vla:2375
        Call vlatracestep(423, vla_step_text(423)) ' vla:2375
    End If
    Call vlafillseries(range("a190:a194"), 2, 2, "growth")
    vla_step = 424 ' vla:2378
    If vlatraceon() Then ' vla:2379
        Call vlatracestep(424, vla_step_text(424)) ' vla:2379
    End If
    Call vlafillseries(range("a200:a203"), 2, 3, "growth")
    vla_step = 425 ' vla:2382
    If vlatraceon() Then ' vla:2383
        Call vlatracestep(425, vla_step_text(425)) ' vla:2383
    End If
    range("z500").font.bold = True
    vla_step = 426 ' vla:2386
    If vlatraceon() Then ' vla:2387
        Call vlatracestep(426, vla_step_text(426)) ' vla:2387
    End If
    Call range("z500").clear
    vla_step = 427 ' vla:2390
    If vlatraceon() Then ' vla:2391
        Call vlatracestep(427, vla_step_text(427)) ' vla:2391
    End If
    Call vlatrimsheet
    vla_step = 428 ' vla:2394
    If vlatraceon() Then ' vla:2395
        Call vlatracestep(428, vla_step_text(428)) ' vla:2395
    End If
    On Error GoTo vla_tryf_5 ' vla:2397 src:792
    vla_step = 429 ' vla:2398 src:792
    If vlatraceon() Then ' vla:2399 src:792
        Call vlatracestep(429, vla_step_text(429)) ' vla:2399 src:792
    End If
    application.displayalerts = False
    Call worksheets("gformat").delete
    application.displayalerts = True
    GoTo vla_tryd_5 ' vla:2402 src:792
vla_tryf_5: ' vla:2403 src:792
    vla_problem = err.description ' vla:2404 src:792
    Resume vla_tryr_5 ' vla:2405 src:792
vla_tryr_5: ' vla:2406 src:792
    On Error GoTo vla_fail ' vla:2407 src:792
vla_tryd_5: ' vla:2408 src:792
    On Error GoTo vla_fail ' vla:2409 src:792
    vla_step = 430 ' vla:2410
    If vlatraceon() Then ' vla:2411
        Call vlatracestep(430, vla_step_text(430)) ' vla:2411
    End If
    Call vlaensuresheet("gformat") ' vla:2413 src:795
    Call worksheets("gformat").activate
    vla_step = 431 ' vla:2414
    If vlatraceon() Then ' vla:2415
        Call vlatracestep(431, vla_step_text(431)) ' vla:2415
    End If
    range("a1:c1").font.bold = True
    vla_step = 432 ' vla:2418
    If vlatraceon() Then ' vla:2419
        Call vlatracestep(432, vla_step_text(432)) ' vla:2419
    End If
    range("a2:c2").font.italic = True
    vla_step = 433 ' vla:2422
    If vlatraceon() Then ' vla:2423
        Call vlatracestep(433, vla_step_text(433)) ' vla:2423
    End If
    range("a3:c3").interior.color = vbblue
    vla_step = 434 ' vla:2426
    If vlatraceon() Then ' vla:2427
        Call vlatracestep(434, vla_step_text(434)) ' vla:2427
    End If
    range("a4:c4").font.color = vlacolor("#FF0000")
    vla_step = 435 ' vla:2430
    If vlatraceon() Then ' vla:2431
        Call vlatracestep(435, vla_step_text(435)) ' vla:2431
    End If
    range("a5:c5").interior.color = vlacolor("#00FF00")
    vla_step = 436 ' vla:2434
    If vlatraceon() Then ' vla:2435
        Call vlatracestep(436, vla_step_text(436)) ' vla:2435
    End If
    range("a6:c6").interior.color = vbyellow
    vla_step = 437 ' vla:2438
    If vlatraceon() Then ' vla:2439
        Call vlatracestep(437, vla_step_text(437)) ' vla:2439
    End If
    range("b6:c6").interior.colorindex = xlnone
    vla_step = 438 ' vla:2442
    If vlatraceon() Then ' vla:2443
        Call vlatracestep(438, vla_step_text(438)) ' vla:2443
    End If
    range("a7:c7").font.size = 16
    vla_step = 439 ' vla:2446
    If vlatraceon() Then ' vla:2447
        Call vlatracestep(439, vla_step_text(439)) ' vla:2447
    End If
    range("a8:c8").font.bold = True
    vla_step = 440 ' vla:2450
    If vlatraceon() Then ' vla:2451
        Call vlatracestep(440, vla_step_text(440)) ' vla:2451
    End If
    range("b8:c8").font.bold = False
    vla_step = 441 ' vla:2454
    If vlatraceon() Then ' vla:2455
        Call vlatracestep(441, vla_step_text(441)) ' vla:2455
    End If
    range("a9:c9").font.italic = True
    vla_step = 442 ' vla:2458
    If vlatraceon() Then ' vla:2459
        Call vlatracestep(442, vla_step_text(442)) ' vla:2459
    End If
    range("c9").font.italic = False
    vla_step = 443 ' vla:2462
    If vlatraceon() Then ' vla:2463
        Call vlatracestep(443, vla_step_text(443)) ' vla:2463
    End If
    range("a10:c10").font.underline = xlunderlinestylesingle
    vla_step = 444 ' vla:2466
    If vlatraceon() Then ' vla:2467
        Call vlatracestep(444, vla_step_text(444)) ' vla:2467
    End If
    range("a11:c11").font.strikethrough = True
    vla_step = 445 ' vla:2470
    If vlatraceon() Then ' vla:2471
        Call vlatracestep(445, vla_step_text(445)) ' vla:2471
    End If
    range("a12:c12").font.underline = xlunderlinestylesingle
    vla_step = 446 ' vla:2474
    If vlatraceon() Then ' vla:2475
        Call vlatracestep(446, vla_step_text(446)) ' vla:2475
    End If
    range("b12:c12").font.underline = xlunderlinestylenone
    vla_step = 447 ' vla:2478
    If vlatraceon() Then ' vla:2479
        Call vlatracestep(447, vla_step_text(447)) ' vla:2479
    End If
    range("a13:c13").font.strikethrough = True
    vla_step = 448 ' vla:2482
    If vlatraceon() Then ' vla:2483
        Call vlatracestep(448, vla_step_text(448)) ' vla:2483
    End If
    range("c13").font.strikethrough = False
    vla_step = 449 ' vla:2486
    If vlatraceon() Then ' vla:2487
        Call vlatracestep(449, vla_step_text(449)) ' vla:2487
    End If
    range("a14:c14").font.name = "Courier New"
    vla_step = 450 ' vla:2490
    If vlatraceon() Then ' vla:2491
        Call vlatracestep(450, vla_step_text(450)) ' vla:2491
    End If
    range("a15:c15").verticalalignment = xltop
    vla_step = 451 ' vla:2494
    If vlatraceon() Then ' vla:2495
        Call vlatracestep(451, vla_step_text(451)) ' vla:2495
    End If
    range("a16:c16").verticalalignment = xlcenter
    vla_step = 452 ' vla:2498
    If vlatraceon() Then ' vla:2499
        Call vlatracestep(452, vla_step_text(452)) ' vla:2499
    End If
    range("a17:c17").verticalalignment = xltop
    vla_step = 453 ' vla:2502
    If vlatraceon() Then ' vla:2503
        Call vlatracestep(453, vla_step_text(453)) ' vla:2503
    End If
    range("b17:c17").verticalalignment = xlbottom
    vla_step = 454 ' vla:2506
    If vlatraceon() Then ' vla:2507
        Call vlatracestep(454, vla_step_text(454)) ' vla:2507
    End If
    range("a18:c18").indentlevel = 2
    vla_step = 455 ' vla:2510
    If vlatraceon() Then ' vla:2511
        Call vlatracestep(455, vla_step_text(455)) ' vla:2511
    End If
    range("a19:c19").orientation = 45
    vla_step = 456 ' vla:2514
    If vlatraceon() Then ' vla:2515
        Call vlatracestep(456, vla_step_text(456)) ' vla:2515
    End If
    range("b21:d23").borders(xledgetop).linestyle = xlcontinuous
    range("b21:d23").borders(xledgebottom).linestyle = xlcontinuous
    range("b21:d23").borders(xledgeleft).linestyle = xlcontinuous
    range("b21:d23").borders(xledgeright).linestyle = xlcontinuous
    vla_step = 457 ' vla:2518
    If vlatraceon() Then ' vla:2519
        Call vlatracestep(457, vla_step_text(457)) ' vla:2519
    End If
    range("b25:d25").borders(xledgebottom).linestyle = xlcontinuous
    vla_step = 458 ' vla:2522
    If vlatraceon() Then ' vla:2523
        Call vlatracestep(458, vla_step_text(458)) ' vla:2523
    End If
    range("b27:d29").borders.linestyle = xlcontinuous
    vla_step = 459 ' vla:2526
    If vlatraceon() Then ' vla:2527
        Call vlatracestep(459, vla_step_text(459)) ' vla:2527
    End If
    range("b31:d33").borders.linestyle = xlcontinuous
    vla_step = 460 ' vla:2530
    If vlatraceon() Then ' vla:2531
        Call vlatracestep(460, vla_step_text(460)) ' vla:2531
    End If
    range("c31:d33").borders.linestyle = xlnone
    vla_step = 461 ' vla:2534
    If vlatraceon() Then ' vla:2535
        Call vlatracestep(461, vla_step_text(461)) ' vla:2535
    End If
    columns("f").columnwidth = 40
    vla_step = 462 ' vla:2538
    If vlatraceon() Then ' vla:2539
        Call vlatracestep(462, vla_step_text(462)) ' vla:2539
    End If
    range("f40") = 1234.56 ' vla:2541 src:847
    vla_step = 463 ' vla:2542
    If vlatraceon() Then ' vla:2543
        Call vlatracestep(463, vla_step_text(463)) ' vla:2543
    End If
    range("f40").numberformat = vlanumberformatcode("number", 2)
    vla_step = 464 ' vla:2546
    If vlatraceon() Then ' vla:2547
        Call vlatracestep(464, vla_step_text(464)) ' vla:2547
    End If
    range("f41") = 1234.56 ' vla:2549 src:849
    vla_step = 465 ' vla:2550
    If vlatraceon() Then ' vla:2551
        Call vlatracestep(465, vla_step_text(465)) ' vla:2551
    End If
    range("f41").numberformat = vlanumberformatcode("number", 3)
    vla_step = 466 ' vla:2554
    If vlatraceon() Then ' vla:2555
        Call vlatracestep(466, vla_step_text(466)) ' vla:2555
    End If
    range("f42") = 1234.56 ' vla:2557 src:851
    vla_step = 467 ' vla:2558
    If vlatraceon() Then ' vla:2559
        Call vlatracestep(467, vla_step_text(467)) ' vla:2559
    End If
    range("f42").numberformat = vlanumberformatcode("number-separated", 2)
    vla_step = 468 ' vla:2562
    If vlatraceon() Then ' vla:2563
        Call vlatracestep(468, vla_step_text(468)) ' vla:2563
    End If
    range("f43") = 1234.56 ' vla:2565 src:853
    vla_step = 469 ' vla:2566
    If vlatraceon() Then ' vla:2567
        Call vlatracestep(469, vla_step_text(469)) ' vla:2567
    End If
    range("f43").numberformat = vlanumberformatcode("number-separated", 0)
    vla_step = 470 ' vla:2570
    If vlatraceon() Then ' vla:2571
        Call vlatracestep(470, vla_step_text(470)) ' vla:2571
    End If
    range("f44") = 1234.56 ' vla:2573 src:855
    vla_step = 471 ' vla:2574
    If vlatraceon() Then ' vla:2575
        Call vlatracestep(471, vla_step_text(471)) ' vla:2575
    End If
    range("f44").numberformat = vlanumberformatcode("dollars", 2)
    vla_step = 472 ' vla:2578
    If vlatraceon() Then ' vla:2579
        Call vlatracestep(472, vla_step_text(472)) ' vla:2579
    End If
    range("f45") = 1234.56 ' vla:2581 src:857
    vla_step = 473 ' vla:2582
    If vlatraceon() Then ' vla:2583
        Call vlatracestep(473, vla_step_text(473)) ' vla:2583
    End If
    range("f45").numberformat = vlanumberformatcode("euros", 0)
    vla_step = 474 ' vla:2586
    If vlatraceon() Then ' vla:2587
        Call vlatracestep(474, vla_step_text(474)) ' vla:2587
    End If
    range("f46") = 1234.56 ' vla:2589 src:859
    vla_step = 475 ' vla:2590
    If vlatraceon() Then ' vla:2591
        Call vlatracestep(475, vla_step_text(475)) ' vla:2591
    End If
    range("f46").numberformat = vlanumberformatcode("pounds", 2)
    vla_step = 476 ' vla:2594
    If vlatraceon() Then ' vla:2595
        Call vlatracestep(476, vla_step_text(476)) ' vla:2595
    End If
    range("f47") = 1234.56 ' vla:2597 src:861
    vla_step = 477 ' vla:2598
    If vlatraceon() Then ' vla:2599
        Call vlatracestep(477, vla_step_text(477)) ' vla:2599
    End If
    range("f47").numberformat = vlanumberformatcode("accounting-dollars", 2)
    vla_step = 478 ' vla:2602
    If vlatraceon() Then ' vla:2603
        Call vlatracestep(478, vla_step_text(478)) ' vla:2603
    End If
    range("f48") = -1234.56 ' vla:2605 src:863
    vla_step = 479 ' vla:2606
    If vlatraceon() Then ' vla:2607
        Call vlatracestep(479, vla_step_text(479)) ' vla:2607
    End If
    range("f48").numberformat = vlanumberformatcode("accounting-euros", 0)
    vla_step = 480 ' vla:2610
    If vlatraceon() Then ' vla:2611
        Call vlatracestep(480, vla_step_text(480)) ' vla:2611
    End If
    range("f49") = 0.125 ' vla:2613 src:865
    vla_step = 481 ' vla:2614
    If vlatraceon() Then ' vla:2615
        Call vlatracestep(481, vla_step_text(481)) ' vla:2615
    End If
    range("f49").numberformat = "0.0%"
    vla_step = 482 ' vla:2618
    If vlatraceon() Then ' vla:2619
        Call vlatracestep(482, vla_step_text(482)) ' vla:2619
    End If
    range("f50") = 0.125 ' vla:2621 src:867
    vla_step = 483 ' vla:2622
    If vlatraceon() Then ' vla:2623
        Call vlatracestep(483, vla_step_text(483)) ' vla:2623
    End If
    range("f50").numberformat = vlanumberformatcode("percent", 2)
    vla_step = 484 ' vla:2626
    If vlatraceon() Then ' vla:2627
        Call vlatracestep(484, vla_step_text(484)) ' vla:2627
    End If
    range("f51") = 46000 ' vla:2629 src:869
    vla_step = 485 ' vla:2630
    If vlatraceon() Then ' vla:2631
        Call vlatracestep(485, vla_step_text(485)) ' vla:2631
    End If
    range("f51").numberformat = "m/d/yyyy"
    vla_step = 486 ' vla:2634
    If vlatraceon() Then ' vla:2635
        Call vlatracestep(486, vla_step_text(486)) ' vla:2635
    End If
    range("f52") = 46000 ' vla:2637 src:871
    vla_step = 487 ' vla:2638
    If vlatraceon() Then ' vla:2639
        Call vlatracestep(487, vla_step_text(487)) ' vla:2639
    End If
    range("f52").numberformat = "[$-F800]dddd, mmmm dd, yyyy"
    vla_step = 488 ' vla:2642
    If vlatraceon() Then ' vla:2643
        Call vlatracestep(488, vla_step_text(488)) ' vla:2643
    End If
    range("f53") = 46000 ' vla:2645 src:873
    vla_step = 489 ' vla:2646
    If vlatraceon() Then ' vla:2647
        Call vlatracestep(489, vla_step_text(489)) ' vla:2647
    End If
    range("f53").numberformat = "yyyy-mm-dd"
    vla_step = 490 ' vla:2650
    If vlatraceon() Then ' vla:2651
        Call vlatracestep(490, vla_step_text(490)) ' vla:2651
    End If
    range("f54") = 0.5625 ' vla:2653 src:875
    vla_step = 491 ' vla:2654
    If vlatraceon() Then ' vla:2655
        Call vlatracestep(491, vla_step_text(491)) ' vla:2655
    End If
    range("f54").numberformat = "[$-F400]h:mm:ss AM/PM"
    vla_step = 492 ' vla:2658
    If vlatraceon() Then ' vla:2659
        Call vlatracestep(492, vla_step_text(492)) ' vla:2659
    End If
    range("f55") = 42 ' vla:2661 src:877
    vla_step = 493 ' vla:2662
    If vlatraceon() Then ' vla:2663
        Call vlatracestep(493, vla_step_text(493)) ' vla:2663
    End If
    range("f55").numberformat = "@"
    vla_step = 494 ' vla:2666
    If vlatraceon() Then ' vla:2667
        Call vlatracestep(494, vla_step_text(494)) ' vla:2667
    End If
    range("f56") = 1234.56 ' vla:2669 src:879
    vla_step = 495 ' vla:2670
    If vlatraceon() Then ' vla:2671
        Call vlatracestep(495, vla_step_text(495)) ' vla:2671
    End If
    range("f56").numberformat = vlanumberformatcode("dollars", 2)
    vla_step = 496 ' vla:2674
    If vlatraceon() Then ' vla:2675
        Call vlatracestep(496, vla_step_text(496)) ' vla:2675
    End If
    range("f56").numberformat = "General"
    vla_step = 497 ' vla:2678
    If vlatraceon() Then ' vla:2679
        Call vlatracestep(497, vla_step_text(497)) ' vla:2679
    End If
    range("f57") = 42 ' vla:2681 src:882
    vla_step = 498 ' vla:2682
    If vlatraceon() Then ' vla:2683
        Call vlatracestep(498, vla_step_text(498)) ' vla:2683
    End If
    range("f57").numberformat = "00000"
    vla_step = 499 ' vla:2686
    If vlatraceon() Then ' vla:2687
        Call vlatracestep(499, vla_step_text(499)) ' vla:2687
    End If
    range("h40:j42").borders(xledgetop).linestyle = xlcontinuous
    range("h40:j42").borders(xledgetop).color = vlacolor("#FF0000")
    range("h40:j42").borders(xledgebottom).linestyle = xlcontinuous
    range("h40:j42").borders(xledgebottom).color = vlacolor("#FF0000")
    range("h40:j42").borders(xledgeleft).linestyle = xlcontinuous
    range("h40:j42").borders(xledgeleft).color = vlacolor("#FF0000")
    range("h40:j42").borders(xledgeright).linestyle = xlcontinuous
    range("h40:j42").borders(xledgeright).color = vlacolor("#FF0000")
    vla_step = 500 ' vla:2690
    If vlatraceon() Then ' vla:2691
        Call vlatracestep(500, vla_step_text(500)) ' vla:2691
    End If
    range("h44:j44").borders(xledgebottom).linestyle = xlcontinuous
    range("h44:j44").borders(xledgebottom).color = vlacolor("green")
    vla_step = 501 ' vla:2694
    If vlatraceon() Then ' vla:2695
        Call vlatracestep(501, vla_step_text(501)) ' vla:2695
    End If
    range("h46:j48").borders.linestyle = xlcontinuous
    range("h46:j48").borders.color = vlacolor("blue")
    vla_step = 502 ' vla:2698
    If vlatraceon() Then ' vla:2699
        Call vlatracestep(502, vla_step_text(502)) ' vla:2699
    End If
    On Error GoTo vla_tryf_6 ' vla:2701 src:899
    vla_step = 503 ' vla:2702 src:899
    If vlatraceon() Then ' vla:2703 src:899
        Call vlatracestep(503, vla_step_text(503)) ' vla:2703 src:899
    End If
    application.displayalerts = False
    Call worksheets("gsortfilter").delete
    application.displayalerts = True
    GoTo vla_tryd_6 ' vla:2706 src:899
vla_tryf_6: ' vla:2707 src:899
    vla_problem = err.description ' vla:2708 src:899
    Resume vla_tryr_6 ' vla:2709 src:899
vla_tryr_6: ' vla:2710 src:899
    On Error GoTo vla_fail ' vla:2711 src:899
vla_tryd_6: ' vla:2712 src:899
    On Error GoTo vla_fail ' vla:2713 src:899
    vla_step = 504 ' vla:2714
    If vlatraceon() Then ' vla:2715
        Call vlatracestep(504, vla_step_text(504)) ' vla:2715
    End If
    Call vlaensuresheet("gsortfilter") ' vla:2717 src:902
    Call worksheets("gsortfilter").activate
    vla_step = 505 ' vla:2718
    If vlatraceon() Then ' vla:2719
        Call vlatracestep(505, vla_step_text(505)) ' vla:2719
    End If
    range("a1") = "Item" ' vla:2721 src:906
    vla_step = 506 ' vla:2722
    If vlatraceon() Then ' vla:2723
        Call vlatracestep(506, vla_step_text(506)) ' vla:2723
    End If
    range("b1") = "Qty" ' vla:2725 src:907
    vla_step = 507 ' vla:2726
    If vlatraceon() Then ' vla:2727
        Call vlatracestep(507, vla_step_text(507)) ' vla:2727
    End If
    range("a2") = "a" ' vla:2729 src:908
    vla_step = 508 ' vla:2730
    If vlatraceon() Then ' vla:2731
        Call vlatracestep(508, vla_step_text(508)) ' vla:2731
    End If
    range("b2") = 2 ' vla:2733 src:909
    vla_step = 509 ' vla:2734
    If vlatraceon() Then ' vla:2735
        Call vlatracestep(509, vla_step_text(509)) ' vla:2735
    End If
    range("a3") = "b" ' vla:2737 src:910
    vla_step = 510 ' vla:2738
    If vlatraceon() Then ' vla:2739
        Call vlatracestep(510, vla_step_text(510)) ' vla:2739
    End If
    range("b3") = 9 ' vla:2741 src:911
    vla_step = 511 ' vla:2742
    If vlatraceon() Then ' vla:2743
        Call vlatracestep(511, vla_step_text(511)) ' vla:2743
    End If
    range("a4") = "c" ' vla:2745 src:912
    vla_step = 512 ' vla:2746
    If vlatraceon() Then ' vla:2747
        Call vlatracestep(512, vla_step_text(512)) ' vla:2747
    End If
    range("b4") = 5 ' vla:2749 src:913
    vla_step = 513 ' vla:2750
    If vlatraceon() Then ' vla:2751
        Call vlatracestep(513, vla_step_text(513)) ' vla:2751
    End If
    Call activesheet.usedrange.sort(key1:=activesheet.usedrange.columns(vlacolumninrange(activesheet.usedrange, "b")), order1:=xldescending, header:=xlyes)
    vla_step = 514 ' vla:2754
    If vlatraceon() Then ' vla:2755
        Call vlatracestep(514, vla_step_text(514)) ' vla:2755
    End If
    range("d1") = "Name" ' vla:2757 src:918
    vla_step = 515 ' vla:2758
    If vlatraceon() Then ' vla:2759
        Call vlatracestep(515, vla_step_text(515)) ' vla:2759
    End If
    range("e1") = "Score" ' vla:2761 src:919
    vla_step = 516 ' vla:2762
    If vlatraceon() Then ' vla:2763
        Call vlatracestep(516, vla_step_text(516)) ' vla:2763
    End If
    range("d2") = "p" ' vla:2765 src:920
    vla_step = 517 ' vla:2766
    If vlatraceon() Then ' vla:2767
        Call vlatracestep(517, vla_step_text(517)) ' vla:2767
    End If
    range("e2") = 3 ' vla:2769 src:921
    vla_step = 518 ' vla:2770
    If vlatraceon() Then ' vla:2771
        Call vlatracestep(518, vla_step_text(518)) ' vla:2771
    End If
    range("d3") = "q" ' vla:2773 src:922
    vla_step = 519 ' vla:2774
    If vlatraceon() Then ' vla:2775
        Call vlatracestep(519, vla_step_text(519)) ' vla:2775
    End If
    range("e3") = 1 ' vla:2777 src:923
    vla_step = 520 ' vla:2778
    If vlatraceon() Then ' vla:2779
        Call vlatracestep(520, vla_step_text(520)) ' vla:2779
    End If
    range("d4") = "r" ' vla:2781 src:924
    vla_step = 521 ' vla:2782
    If vlatraceon() Then ' vla:2783
        Call vlatracestep(521, vla_step_text(521)) ' vla:2783
    End If
    range("e4") = 2 ' vla:2785 src:925
    vla_step = 522 ' vla:2786
    If vlatraceon() Then ' vla:2787
        Call vlatracestep(522, vla_step_text(522)) ' vla:2787
    End If
    Call range("d1:e4").sort(key1:=range("d1:e4").columns(vlacolumninrange(range("d1:e4"), "e")), order1:=xlascending, header:=xlyes)
    vla_step = 523 ' vla:2790
    If vlatraceon() Then ' vla:2791
        Call vlatracestep(523, vla_step_text(523)) ' vla:2791
    End If
    range("g2") = "x" ' vla:2793 src:929
    vla_step = 524 ' vla:2794
    If vlatraceon() Then ' vla:2795
        Call vlatracestep(524, vla_step_text(524)) ' vla:2795
    End If
    range("h2") = 5 ' vla:2797 src:930
    vla_step = 525 ' vla:2798
    If vlatraceon() Then ' vla:2799
        Call vlatracestep(525, vla_step_text(525)) ' vla:2799
    End If
    range("g3") = "y" ' vla:2801 src:931
    vla_step = 526 ' vla:2802
    If vlatraceon() Then ' vla:2803
        Call vlatracestep(526, vla_step_text(526)) ' vla:2803
    End If
    range("h3") = 9 ' vla:2805 src:932
    vla_step = 527 ' vla:2806
    If vlatraceon() Then ' vla:2807
        Call vlatracestep(527, vla_step_text(527)) ' vla:2807
    End If
    range("g4") = "z" ' vla:2809 src:933
    vla_step = 528 ' vla:2810
    If vlatraceon() Then ' vla:2811
        Call vlatracestep(528, vla_step_text(528)) ' vla:2811
    End If
    range("h4") = 7 ' vla:2813 src:934
    vla_step = 529 ' vla:2814
    If vlatraceon() Then ' vla:2815
        Call vlatracestep(529, vla_step_text(529)) ' vla:2815
    End If
    Call range("g2:h4").sort(key1:=range("g2:h4").columns(vlacolumninrange(range("g2:h4"), "h")), order1:=xldescending, header:=xlno)
    vla_step = 530 ' vla:2818
    If vlatraceon() Then ' vla:2819
        Call vlatracestep(530, vla_step_text(530)) ' vla:2819
    End If
    range("j1") = "Region" ' vla:2821 src:939
    vla_step = 531 ' vla:2822
    If vlatraceon() Then ' vla:2823
        Call vlatracestep(531, vla_step_text(531)) ' vla:2823
    End If
    range("k1") = "Amount" ' vla:2825 src:940
    vla_step = 532 ' vla:2826
    If vlatraceon() Then ' vla:2827
        Call vlatracestep(532, vla_step_text(532)) ' vla:2827
    End If
    range("l1") = "Tag" ' vla:2829 src:941
    vla_step = 533 ' vla:2830
    If vlatraceon() Then ' vla:2831
        Call vlatracestep(533, vla_step_text(533)) ' vla:2831
    End If
    range("j2") = "West" ' vla:2833 src:942
    vla_step = 534 ' vla:2834
    If vlatraceon() Then ' vla:2835
        Call vlatracestep(534, vla_step_text(534)) ' vla:2835
    End If
    range("k2") = 100 ' vla:2837 src:943
    vla_step = 535 ' vla:2838
    If vlatraceon() Then ' vla:2839
        Call vlatracestep(535, vla_step_text(535)) ' vla:2839
    End If
    range("l2") = "a" ' vla:2841 src:944
    vla_step = 536 ' vla:2842
    If vlatraceon() Then ' vla:2843
        Call vlatracestep(536, vla_step_text(536)) ' vla:2843
    End If
    range("j3") = "East" ' vla:2845 src:945
    vla_step = 537 ' vla:2846
    If vlatraceon() Then ' vla:2847
        Call vlatracestep(537, vla_step_text(537)) ' vla:2847
    End If
    range("k3") = 50 ' vla:2849 src:946
    vla_step = 538 ' vla:2850
    If vlatraceon() Then ' vla:2851
        Call vlatracestep(538, vla_step_text(538)) ' vla:2851
    End If
    range("l3") = "b" ' vla:2853 src:947
    vla_step = 539 ' vla:2854
    If vlatraceon() Then ' vla:2855
        Call vlatracestep(539, vla_step_text(539)) ' vla:2855
    End If
    range("j4") = "West" ' vla:2857 src:948
    vla_step = 540 ' vla:2858
    If vlatraceon() Then ' vla:2859
        Call vlatracestep(540, vla_step_text(540)) ' vla:2859
    End If
    range("k4") = 300 ' vla:2861 src:949
    vla_step = 541 ' vla:2862
    If vlatraceon() Then ' vla:2863
        Call vlatracestep(541, vla_step_text(541)) ' vla:2863
    End If
    range("l4") = "c" ' vla:2865 src:950
    vla_step = 542 ' vla:2866
    If vlatraceon() Then ' vla:2867
        Call vlatracestep(542, vla_step_text(542)) ' vla:2867
    End If
    range("j5") = "East" ' vla:2869 src:951
    vla_step = 543 ' vla:2870
    If vlatraceon() Then ' vla:2871
        Call vlatracestep(543, vla_step_text(543)) ' vla:2871
    End If
    range("k5") = 250 ' vla:2873 src:952
    vla_step = 544 ' vla:2874
    If vlatraceon() Then ' vla:2875
        Call vlatracestep(544, vla_step_text(544)) ' vla:2875
    End If
    range("l5") = "d" ' vla:2877 src:953
    vla_step = 545 ' vla:2878
    If vlatraceon() Then ' vla:2879
        Call vlatracestep(545, vla_step_text(545)) ' vla:2879
    End If
    Call range("j1:l5").sort(key1:=range("j1:l5").columns(vlacolumninrange(range("j1:l5"), "j")), order1:=xlascending, key2:=range("j1:l5").columns(vlacolumninrange(range("j1:l5"), "k")), order2:=xldescending, header:=xlyes)
    vla_step = 546 ' vla:2882
    If vlatraceon() Then ' vla:2883
        Call vlatracestep(546, vla_step_text(546)) ' vla:2883
    End If
    range("n2") = 2 ' vla:2885 src:957
    vla_step = 547 ' vla:2886
    If vlatraceon() Then ' vla:2887
        Call vlatracestep(547, vla_step_text(547)) ' vla:2887
    End If
    range("o2") = 9 ' vla:2889 src:958
    vla_step = 548 ' vla:2890
    If vlatraceon() Then ' vla:2891
        Call vlatracestep(548, vla_step_text(548)) ' vla:2891
    End If
    range("n3") = 1 ' vla:2893 src:959
    vla_step = 549 ' vla:2894
    If vlatraceon() Then ' vla:2895
        Call vlatracestep(549, vla_step_text(549)) ' vla:2895
    End If
    range("o3") = 8 ' vla:2897 src:960
    vla_step = 550 ' vla:2898
    If vlatraceon() Then ' vla:2899
        Call vlatracestep(550, vla_step_text(550)) ' vla:2899
    End If
    range("n4") = 2 ' vla:2901 src:961
    vla_step = 551 ' vla:2902
    If vlatraceon() Then ' vla:2903
        Call vlatracestep(551, vla_step_text(551)) ' vla:2903
    End If
    range("o4") = 7 ' vla:2905 src:962
    vla_step = 552 ' vla:2906
    If vlatraceon() Then ' vla:2907
        Call vlatracestep(552, vla_step_text(552)) ' vla:2907
    End If
    Call range("n2:o4").sort(key1:=range("n2:o4").columns(vlacolumninrange(range("n2:o4"), "n")), order1:=xlascending, key2:=range("n2:o4").columns(vlacolumninrange(range("n2:o4"), "o")), order2:=xlascending, header:=xlno)
    vla_step = 553 ' vla:2910
    If vlatraceon() Then ' vla:2911
        Call vlatracestep(553, vla_step_text(553)) ' vla:2911
    End If
    range("a40") = "Region" ' vla:2913 src:972
    vla_step = 554 ' vla:2914
    If vlatraceon() Then ' vla:2915
        Call vlatracestep(554, vla_step_text(554)) ' vla:2915
    End If
    range("b40") = "Amount" ' vla:2917 src:973
    vla_step = 555 ' vla:2918
    If vlatraceon() Then ' vla:2919
        Call vlatracestep(555, vla_step_text(555)) ' vla:2919
    End If
    range("a41") = "West" ' vla:2921 src:974
    vla_step = 556 ' vla:2922
    If vlatraceon() Then ' vla:2923
        Call vlatracestep(556, vla_step_text(556)) ' vla:2923
    End If
    range("b41") = 1 ' vla:2925 src:975
    vla_step = 557 ' vla:2926
    If vlatraceon() Then ' vla:2927
        Call vlatracestep(557, vla_step_text(557)) ' vla:2927
    End If
    range("a42") = "East" ' vla:2929 src:976
    vla_step = 558 ' vla:2930
    If vlatraceon() Then ' vla:2931
        Call vlatracestep(558, vla_step_text(558)) ' vla:2931
    End If
    range("b42") = 2 ' vla:2933 src:977
    vla_step = 559 ' vla:2934
    If vlatraceon() Then ' vla:2935
        Call vlatracestep(559, vla_step_text(559)) ' vla:2935
    End If
    range("a43") = "West" ' vla:2937 src:978
    vla_step = 560 ' vla:2938
    If vlatraceon() Then ' vla:2939
        Call vlatracestep(560, vla_step_text(560)) ' vla:2939
    End If
    range("b43") = 3 ' vla:2941 src:979
    vla_step = 561 ' vla:2942
    If vlatraceon() Then ' vla:2943
        Call vlatracestep(561, vla_step_text(561)) ' vla:2943
    End If
    Call range("a40:b43").autofilter(field:=vlafilterfield(range("a40:b43"), "a"), criteria1:=vlafiltercriterion("at-least", "West"), operator:=xland, criteria2:=vlafiltercriterion("at-most", "West"))
    vla_step = 562 ' vla:2946
    If vlatraceon() Then ' vla:2947
        Call vlatracestep(562, vla_step_text(562)) ' vla:2947
    End If
    Call range("a40:b43").specialcells(xlcelltypevisible).copy(destination:=range("d46"))
    vla_step = 563 ' vla:2950
    If vlatraceon() Then ' vla:2951
        Call vlatracestep(563, vla_step_text(563)) ' vla:2951
    End If
    activesheet.autofiltermode = False
    vla_step = 564 ' vla:2954
    If vlatraceon() Then ' vla:2955
        Call vlatracestep(564, vla_step_text(564)) ' vla:2955
    End If
    On Error GoTo vla_tryf_7 ' vla:2957 src:988
    vla_step = 565 ' vla:2958 src:988
    If vlatraceon() Then ' vla:2959 src:988
        Call vlatracestep(565, vla_step_text(565)) ' vla:2959 src:988
    End If
    Call range("a40:b43").autofilter(field:=vlafilterfield(range("a40:b43"), "b"), criteria1:=vlafiltercriterion("greater", "abc"))
    GoTo vla_tryd_7 ' vla:2962 src:988
vla_tryf_7: ' vla:2963 src:988
    vla_problem = err.description ' vla:2964 src:988
    Resume vla_tryr_7 ' vla:2965 src:988
vla_tryr_7: ' vla:2966 src:988
    On Error GoTo vla_fail ' vla:2967 src:988
vla_tryd_7: ' vla:2968 src:988
    On Error GoTo vla_fail ' vla:2969 src:988
    vla_step = 566 ' vla:2970
    If vlatraceon() Then ' vla:2971
        Call vlatracestep(566, vla_step_text(566)) ' vla:2971
    End If
    range("a20") = "Region" ' vla:2973 src:994
    vla_step = 567 ' vla:2974
    If vlatraceon() Then ' vla:2975
        Call vlatracestep(567, vla_step_text(567)) ' vla:2975
    End If
    range("b20") = "Amount" ' vla:2977 src:995
    vla_step = 568 ' vla:2978
    If vlatraceon() Then ' vla:2979
        Call vlatracestep(568, vla_step_text(568)) ' vla:2979
    End If
    range("c20") = "Code" ' vla:2981 src:996
    vla_step = 569 ' vla:2982
    If vlatraceon() Then ' vla:2983
        Call vlatracestep(569, vla_step_text(569)) ' vla:2983
    End If
    range("a21") = "West" ' vla:2985 src:997
    vla_step = 570 ' vla:2986
    If vlatraceon() Then ' vla:2987
        Call vlatracestep(570, vla_step_text(570)) ' vla:2987
    End If
    range("b21") = 100 ' vla:2989 src:998
    vla_step = 571 ' vla:2990
    If vlatraceon() Then ' vla:2991
        Call vlatracestep(571, vla_step_text(571)) ' vla:2991
    End If
    range("c21") = "5*3" ' vla:2993 src:999
    vla_step = 572 ' vla:2994
    If vlatraceon() Then ' vla:2995
        Call vlatracestep(572, vla_step_text(572)) ' vla:2995
    End If
    range("a22") = "East" ' vla:2997 src:1000
    vla_step = 573 ' vla:2998
    If vlatraceon() Then ' vla:2999
        Call vlatracestep(573, vla_step_text(573)) ' vla:2999
    End If
    range("b22") = 250 ' vla:3001 src:1001
    vla_step = 574 ' vla:3002
    If vlatraceon() Then ' vla:3003
        Call vlatracestep(574, vla_step_text(574)) ' vla:3003
    End If
    range("c22") = "53" ' vla:3005 src:1002
    vla_step = 575 ' vla:3006
    If vlatraceon() Then ' vla:3007
        Call vlatracestep(575, vla_step_text(575)) ' vla:3007
    End If
    range("a23") = "Western" ' vla:3009 src:1003
    vla_step = 576 ' vla:3010
    If vlatraceon() Then ' vla:3011
        Call vlatracestep(576, vla_step_text(576)) ' vla:3011
    End If
    range("b23") = 100 ' vla:3013 src:1004
    vla_step = 577 ' vla:3014
    If vlatraceon() Then ' vla:3015
        Call vlatracestep(577, vla_step_text(577)) ' vla:3015
    End If
    range("c23") = "x" ' vla:3017 src:1005
    vla_step = 578 ' vla:3018
    If vlatraceon() Then ' vla:3019
        Call vlatracestep(578, vla_step_text(578)) ' vla:3019
    End If
    range("a24") = "West" ' vla:3021 src:1006
    vla_step = 579 ' vla:3022
    If vlatraceon() Then ' vla:3023
        Call vlatracestep(579, vla_step_text(579)) ' vla:3023
    End If
    range("b24") = 50 ' vla:3025 src:1007
    vla_step = 580 ' vla:3026
    If vlatraceon() Then ' vla:3027
        Call vlatracestep(580, vla_step_text(580)) ' vla:3027
    End If
    range("c24") = "y" ' vla:3029 src:1008
    vla_step = 581 ' vla:3030
    If vlatraceon() Then ' vla:3031
        Call vlatracestep(581, vla_step_text(581)) ' vla:3031
    End If
    range("a25") = "east" ' vla:3033 src:1009
    vla_step = 582 ' vla:3034
    If vlatraceon() Then ' vla:3035
        Call vlatracestep(582, vla_step_text(582)) ' vla:3035
    End If
    range("b25") = 300 ' vla:3037 src:1010
    vla_step = 583 ' vla:3038
    If vlatraceon() Then ' vla:3039
        Call vlatracestep(583, vla_step_text(583)) ' vla:3039
    End If
    range("c25") = "z" ' vla:3041 src:1011
    vla_step = 584 ' vla:3042
    If vlatraceon() Then ' vla:3043
        Call vlatracestep(584, vla_step_text(584)) ' vla:3043
    End If
    range("a26") = "North" ' vla:3045 src:1012
    vla_step = 585 ' vla:3046
    If vlatraceon() Then ' vla:3047
        Call vlatracestep(585, vla_step_text(585)) ' vla:3047
    End If
    range("b26") = 175 ' vla:3049 src:1013
    vla_step = 586 ' vla:3050
    If vlatraceon() Then ' vla:3051
        Call vlatracestep(586, vla_step_text(586)) ' vla:3051
    End If
    range("c26") = "w" ' vla:3053 src:1014
    vla_step = 587 ' vla:3054
    If vlatraceon() Then ' vla:3055
        Call vlatracestep(587, vla_step_text(587)) ' vla:3055
    End If
    range("a27") = "East" ' vla:3057 src:1015
    vla_step = 588 ' vla:3058
    If vlatraceon() Then ' vla:3059
        Call vlatracestep(588, vla_step_text(588)) ' vla:3059
    End If
    range("b27") = 75 ' vla:3061 src:1016
    vla_step = 589 ' vla:3062
    If vlatraceon() Then ' vla:3063
        Call vlatracestep(589, vla_step_text(589)) ' vla:3063
    End If
    range("c27") = "v" ' vla:3065 src:1017
    vla_step = 590 ' vla:3066
    If vlatraceon() Then ' vla:3067
        Call vlatracestep(590, vla_step_text(590)) ' vla:3067
    End If
    range("b21:b27").numberformat = vlanumberformatcode("dollars", 2)
    vla_step = 591 ' vla:3070
    If vlatraceon() Then ' vla:3071
        Call vlatracestep(591, vla_step_text(591)) ' vla:3071
    End If
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "a"), criteria1:=vlafiltercriterion("at-least", "West"), operator:=xland, criteria2:=vlafiltercriterion("at-most", "West"))
    vla_step = 592 ' vla:3074
    If vlatraceon() Then ' vla:3075
        Call vlatracestep(592, vla_step_text(592)) ' vla:3075
    End If
    Call range("a20:c27").specialcells(xlcelltypevisible).copy(destination:=range("e30"))
    vla_step = 593 ' vla:3078
    If vlatraceon() Then ' vla:3079
        Call vlatracestep(593, vla_step_text(593)) ' vla:3079
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 594 ' vla:3082
    If vlatraceon() Then ' vla:3083
        Call vlatracestep(594, vla_step_text(594)) ' vla:3083
    End If
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "b"), criteria1:=vlafiltercriterion("at-least", 100), operator:=xland, criteria2:=vlafiltercriterion("at-most", 100))
    vla_step = 595 ' vla:3086
    If vlatraceon() Then ' vla:3087
        Call vlatracestep(595, vla_step_text(595)) ' vla:3087
    End If
    Call range("a20:c27").specialcells(xlcelltypevisible).copy(destination:=range("i30"))
    vla_step = 596 ' vla:3090
    If vlatraceon() Then ' vla:3091
        Call vlatracestep(596, vla_step_text(596)) ' vla:3091
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 597 ' vla:3094
    If vlatraceon() Then ' vla:3095
        Call vlatracestep(597, vla_step_text(597)) ' vla:3095
    End If
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "c"), criteria1:=vlafiltercriterion("contains", "*"))
    vla_step = 598 ' vla:3098
    If vlatraceon() Then ' vla:3099
        Call vlatracestep(598, vla_step_text(598)) ' vla:3099
    End If
    Call range("a20:c27").specialcells(xlcelltypevisible).copy(destination:=range("m30"))
    vla_step = 599 ' vla:3102
    If vlatraceon() Then ' vla:3103
        Call vlatracestep(599, vla_step_text(599)) ' vla:3103
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 600 ' vla:3106
    If vlatraceon() Then ' vla:3107
        Call vlatracestep(600, vla_step_text(600)) ' vla:3107
    End If
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "b"), criteria1:=vlafiltercriterion("greater", 150))
    vla_step = 601 ' vla:3110
    If vlatraceon() Then ' vla:3111
        Call vlatracestep(601, vla_step_text(601)) ' vla:3111
    End If
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "a"), criteria1:=vlafiltercriterion("at-least", "East"), operator:=xland, criteria2:=vlafiltercriterion("at-most", "East"))
    vla_step = 602 ' vla:3114
    If vlatraceon() Then ' vla:3115
        Call vlatracestep(602, vla_step_text(602)) ' vla:3115
    End If
    Call range("a20:c27").specialcells(xlcelltypevisible).copy(destination:=range("q30"))
    vla_step = 603 ' vla:3118
    If vlatraceon() Then ' vla:3119
        Call vlatracestep(603, vla_step_text(603)) ' vla:3119
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 604 ' vla:3122
    If vlatraceon() Then ' vla:3123
        Call vlatracestep(604, vla_step_text(604)) ' vla:3123
    End If
    Call range("a20:c27").autofilter(field:=vlafilterfield(range("a20:c27"), "b"), criteria1:=vlafiltercriterion("less", 99.5))
    vla_step = 605 ' vla:3126
    If vlatraceon() Then ' vla:3127
        Call vlatracestep(605, vla_step_text(605)) ' vla:3127
    End If
    Call range("a20:c27").specialcells(xlcelltypevisible).copy(destination:=range("u30"))
    vla_step = 606 ' vla:3130
    If vlatraceon() Then ' vla:3131
        Call vlatracestep(606, vla_step_text(606)) ' vla:3131
    End If
    Call vlaaddfilters(range("a20:c27"))
    vla_step = 607 ' vla:3134
    If vlatraceon() Then ' vla:3135
        Call vlatracestep(607, vla_step_text(607)) ' vla:3135
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 608 ' vla:3138
    If vlatraceon() Then ' vla:3139
        Call vlatracestep(608, vla_step_text(608)) ' vla:3139
    End If
    On Error GoTo vla_tryf_8 ' vla:3141 src:1051
    vla_step = 609 ' vla:3142 src:1051
    If vlatraceon() Then ' vla:3143 src:1051
        Call vlatracestep(609, vla_step_text(609)) ' vla:3143 src:1051
    End If
    Call vlaaddfilters(range("a40:b43"))
    GoTo vla_tryd_8 ' vla:3146 src:1051
vla_tryf_8: ' vla:3147 src:1051
    vla_problem = err.description ' vla:3148 src:1051
    Resume vla_tryr_8 ' vla:3149 src:1051
vla_tryr_8: ' vla:3150 src:1051
    On Error GoTo vla_fail ' vla:3151 src:1051
vla_tryd_8: ' vla:3152 src:1051
    On Error GoTo vla_fail ' vla:3153 src:1051
    vla_step = 610 ' vla:3154
    If vlatraceon() Then ' vla:3155
        Call vlatracestep(610, vla_step_text(610)) ' vla:3155
    End If
    On Error GoTo vla_tryf_9 ' vla:3157 src:1060
    vla_step = 611 ' vla:3158 src:1060
    If vlatraceon() Then ' vla:3159 src:1060
        Call vlatracestep(611, vla_step_text(611)) ' vla:3159 src:1060
    End If
    application.displayalerts = False
    Call worksheets("gtext").delete
    application.displayalerts = True
    GoTo vla_tryd_9 ' vla:3162 src:1060
vla_tryf_9: ' vla:3163 src:1060
    vla_problem = err.description ' vla:3164 src:1060
    Resume vla_tryr_9 ' vla:3165 src:1060
vla_tryr_9: ' vla:3166 src:1060
    On Error GoTo vla_fail ' vla:3167 src:1060
vla_tryd_9: ' vla:3168 src:1060
    On Error GoTo vla_fail ' vla:3169 src:1060
    vla_step = 612 ' vla:3170
    If vlatraceon() Then ' vla:3171
        Call vlatracestep(612, vla_step_text(612)) ' vla:3171
    End If
    Call vlaensuresheet("gtext") ' vla:3173 src:1063
    Call worksheets("gtext").activate
    vla_step = 613 ' vla:3174
    If vlatraceon() Then ' vla:3175
        Call vlatracestep(613, vla_step_text(613)) ' vla:3175
    End If
    range("a1") = "widget" ' vla:3177 src:1067
    vla_step = 614 ' vla:3178
    If vlatraceon() Then ' vla:3179
        Call vlatracestep(614, vla_step_text(614)) ' vla:3179
    End If
    range("a2") = 42 ' vla:3181 src:1068
    vla_step = 615 ' vla:3182
    If vlatraceon() Then ' vla:3183
        Call vlatracestep(615, vla_step_text(615)) ' vla:3183
    End If
    range("a3").Formula2 = "=CHAR(97)&CHAR(98)"
    vla_step = 616 ' vla:3186
    If vlatraceon() Then ' vla:3187
        Call vlatracestep(616, vla_step_text(616)) ' vla:3187
    End If
    range("a4") = "'true" ' vla:3189 src:1070
    vla_step = 617 ' vla:3190
    If vlatraceon() Then ' vla:3191
        Call vlatracestep(617, vla_step_text(617)) ' vla:3191
    End If
    range("a5") = "'=abc" ' vla:3193 src:1071
    vla_step = 618 ' vla:3194
    If vlatraceon() Then ' vla:3195
        Call vlatracestep(618, vla_step_text(618)) ' vla:3195
    End If
    Call vlatextinrange(range("a1:a5"), "upper")
    vla_step = 619 ' vla:3198
    If vlatraceon() Then ' vla:3199
        Call vlatracestep(619, vla_step_text(619)) ' vla:3199
    End If
    range("b1") = "HELLO World" ' vla:3201 src:1075
    vla_step = 620 ' vla:3202
    If vlatraceon() Then ' vla:3203
        Call vlatracestep(620, vla_step_text(620)) ' vla:3203
    End If
    range("b2") = "MiXeD" ' vla:3205 src:1076
    vla_step = 621 ' vla:3206
    If vlatraceon() Then ' vla:3207
        Call vlatracestep(621, vla_step_text(621)) ' vla:3207
    End If
    Call vlatextinrange(columns("b"), "lower")
    vla_step = 622 ' vla:3210
    If vlatraceon() Then ' vla:3211
        Call vlatracestep(622, vla_step_text(622)) ' vla:3211
    End If
    range("c1") = "don't stop" ' vla:3213 src:1080
    vla_step = 623 ' vla:3214
    If vlatraceon() Then ' vla:3215
        Call vlatracestep(623, vla_step_text(623)) ' vla:3215
    End If
    range("c2") = "3rd quarter" ' vla:3217 src:1081
    vla_step = 624 ' vla:3218
    If vlatraceon() Then ' vla:3219
        Call vlatracestep(624, vla_step_text(624)) ' vla:3219
    End If
    range("c3") = "o'neil" ' vla:3221 src:1082
    vla_step = 625 ' vla:3222
    If vlatraceon() Then ' vla:3223
        Call vlatracestep(625, vla_step_text(625)) ' vla:3223
    End If
    Call vlatextinrange(range("c1:c3"), "capitalize-after-space")
    vla_step = 626 ' vla:3226
    If vlatraceon() Then ' vla:3227
        Call vlatracestep(626, vla_step_text(626)) ' vla:3227
    End If
    range("d1") = "don't stop" ' vla:3229 src:1084
    vla_step = 627 ' vla:3230
    If vlatraceon() Then ' vla:3231
        Call vlatracestep(627, vla_step_text(627)) ' vla:3231
    End If
    range("d2") = "3rd quarter" ' vla:3233 src:1085
    vla_step = 628 ' vla:3234
    If vlatraceon() Then ' vla:3235
        Call vlatracestep(628, vla_step_text(628)) ' vla:3235
    End If
    range("d3") = "o'neil" ' vla:3237 src:1086
    vla_step = 629 ' vla:3238
    If vlatraceon() Then ' vla:3239
        Call vlatracestep(629, vla_step_text(629)) ' vla:3239
    End If
    Call vlatextinrange(columns("d"), "capitalize-after-non-letter")
    vla_step = 630 ' vla:3242
    If vlatraceon() Then ' vla:3243
        Call vlatracestep(630, vla_step_text(630)) ' vla:3243
    End If
    range("e1") = "  a   b  " ' vla:3245 src:1092
    vla_step = 631 ' vla:3246
    If vlatraceon() Then ' vla:3247
        Call vlatracestep(631, vla_step_text(631)) ' vla:3247
    End If
    range("e2") = "'  00123 " ' vla:3249 src:1093
    vla_step = 632 ' vla:3250
    If vlatraceon() Then ' vla:3251
        Call vlatracestep(632, vla_step_text(632)) ' vla:3251
    End If
    range("e3") = "   " ' vla:3253 src:1094
    vla_step = 633 ' vla:3254
    If vlatraceon() Then ' vla:3255
        Call vlatracestep(633, vla_step_text(633)) ' vla:3255
    End If
    range("e4").Formula2 = "=UNICHAR(160)&CHAR(120)&UNICHAR(160)&UNICHAR(160)&CHAR(121)"
    vla_step = 634 ' vla:3258
    If vlatraceon() Then ' vla:3259
        Call vlatracestep(634, vla_step_text(634)) ' vla:3259
    End If
    Call range("e4:e4").copy
    Call range("e4:e4").pastespecial(paste:=xlpastevalues)
    application.cutcopymode = False
    vla_step = 635 ' vla:3262
    If vlatraceon() Then ' vla:3263
        Call vlatracestep(635, vla_step_text(635)) ' vla:3263
    End If
    Call vlatextinrange(range("e1:e4"), "remove-extra-spaces")
    vla_step = 636 ' vla:3266
    If vlatraceon() Then ' vla:3267
        Call vlatracestep(636, vla_step_text(636)) ' vla:3267
    End If
    range("f1").Formula2 = "=CHAR(97)&CHAR(10)&CHAR(98)"
    vla_step = 637 ' vla:3270
    If vlatraceon() Then ' vla:3271
        Call vlatracestep(637, vla_step_text(637)) ' vla:3271
    End If
    range("f2").Formula2 = "=CHAR(99)&CHAR(9)&CHAR(100)"
    vla_step = 638 ' vla:3274
    If vlatraceon() Then ' vla:3275
        Call vlatracestep(638, vla_step_text(638)) ' vla:3275
    End If
    Call range("f1:f2").copy
    Call range("f1:f2").pastespecial(paste:=xlpastevalues)
    application.cutcopymode = False
    vla_step = 639 ' vla:3278
    If vlatraceon() Then ' vla:3279
        Call vlatracestep(639, vla_step_text(639)) ' vla:3279
    End If
    Call vlatextinrange(range("f1:f2"), "remove-non-printing")
    vla_step = 640 ' vla:3282
    If vlatraceon() Then ' vla:3283
        Call vlatracestep(640, vla_step_text(640)) ' vla:3283
    End If
    range("g1") = "Widget" ' vla:3285 src:1108
    vla_step = 641 ' vla:3286
    If vlatraceon() Then ' vla:3287
        Call vlatracestep(641, vla_step_text(641)) ' vla:3287
    End If
    range("g2") = "banana" ' vla:3289 src:1109
    vla_step = 642 ' vla:3290
    If vlatraceon() Then ' vla:3291
        Call vlatracestep(642, vla_step_text(642)) ' vla:3291
    End If
    gtext_row = vlafindrow("Widget", "g") ' vla:3293 src:1110
    vla_step = 643 ' vla:3294
    If vlatraceon() Then ' vla:3295
        Call vlatracestep(643, vla_step_text(643)) ' vla:3295
    End If
    Call vlareplaceinrange(range("g2:g2"), "AN", "xy", "values")
    vla_step = 644 ' vla:3298
    If vlatraceon() Then ' vla:3299
        Call vlatracestep(644, vla_step_text(644)) ' vla:3299
    End If
    range("i1:i20").numberformat = "@"
    vla_step = 645 ' vla:3302
    If vlatraceon() Then ' vla:3303
        Call vlatracestep(645, vla_step_text(645)) ' vla:3303
    End If
    gtext_code = "INV-ab-Cd" ' vla:3305 src:1116
    vla_step = 646 ' vla:3306
    If vlatraceon() Then ' vla:3307
        Call vlatracestep(646, vla_step_text(646)) ' vla:3307
    End If
    gtext_part = vlatextbeside(gtext_code, "-", "before", "first") ' vla:3309 src:1117
    vla_step = 647 ' vla:3310
    If vlatraceon() Then ' vla:3311
        Call vlatracestep(647, vla_step_text(647)) ' vla:3311
    End If
    range("i1") = gtext_part ' vla:3313 src:1118
    vla_step = 648 ' vla:3314
    If vlatraceon() Then ' vla:3315
        Call vlatracestep(648, vla_step_text(648)) ' vla:3315
    End If
    gtext_part = vlatextbeside(gtext_code, "-", "after", "first") ' vla:3317 src:1119
    vla_step = 649 ' vla:3318
    If vlatraceon() Then ' vla:3319
        Call vlatracestep(649, vla_step_text(649)) ' vla:3319
    End If
    range("i2") = gtext_part ' vla:3321 src:1120
    vla_step = 650 ' vla:3322
    If vlatraceon() Then ' vla:3323
        Call vlatracestep(650, vla_step_text(650)) ' vla:3323
    End If
    gtext_part = vlatextbeside(gtext_code, "-", "after", "last") ' vla:3325 src:1121
    vla_step = 651 ' vla:3326
    If vlatraceon() Then ' vla:3327
        Call vlatracestep(651, vla_step_text(651)) ' vla:3327
    End If
    range("i3") = gtext_part ' vla:3329 src:1122
    vla_step = 652 ' vla:3330
    If vlatraceon() Then ' vla:3331
        Call vlatracestep(652, vla_step_text(652)) ' vla:3331
    End If
    gtext_part = vlatextbeside(gtext_code, "c", "before", "first") ' vla:3333 src:1125
    vla_step = 653 ' vla:3334
    If vlatraceon() Then ' vla:3335
        Call vlatracestep(653, vla_step_text(653)) ' vla:3335
    End If
    range("i4") = gtext_part ' vla:3337 src:1126
    vla_step = 654 ' vla:3338
    If vlatraceon() Then ' vla:3339
        Call vlatracestep(654, vla_step_text(654)) ' vla:3339
    End If
    gtext_part = left(gtext_code, 3) ' vla:3341 src:1127
    vla_step = 655 ' vla:3342
    If vlatraceon() Then ' vla:3343
        Call vlatracestep(655, vla_step_text(655)) ' vla:3343
    End If
    range("i5") = gtext_part ' vla:3345 src:1128
    vla_step = 656 ' vla:3346
    If vlatraceon() Then ' vla:3347
        Call vlatracestep(656, vla_step_text(656)) ' vla:3347
    End If
    gtext_part = right(gtext_code, 2) ' vla:3349 src:1129
    vla_step = 657 ' vla:3350
    If vlatraceon() Then ' vla:3351
        Call vlatracestep(657, vla_step_text(657)) ' vla:3351
    End If
    range("i6") = gtext_part ' vla:3353 src:1130
    vla_step = 658 ' vla:3354
    If vlatraceon() Then ' vla:3355
        Call vlatracestep(658, vla_step_text(658)) ' vla:3355
    End If
    gtext_part = vlatextpad(42, "0", 5, "left") ' vla:3357 src:1132
    vla_step = 659 ' vla:3358
    If vlatraceon() Then ' vla:3359
        Call vlatracestep(659, vla_step_text(659)) ' vla:3359
    End If
    range("i7") = gtext_part ' vla:3361 src:1133
    vla_step = 660 ' vla:3362
    If vlatraceon() Then ' vla:3363
        Call vlatracestep(660, vla_step_text(660)) ' vla:3363
    End If
    gtext_part = vlatextpad("ab", ".", 4, "right") ' vla:3365 src:1134
    vla_step = 661 ' vla:3366
    If vlatraceon() Then ' vla:3367
        Call vlatracestep(661, vla_step_text(661)) ' vla:3367
    End If
    range("i8") = gtext_part ' vla:3369 src:1135
    vla_step = 662 ' vla:3370
    If vlatraceon() Then ' vla:3371
        Call vlatracestep(662, vla_step_text(662)) ' vla:3371
    End If
    range("j1") = "x" ' vla:3373 src:1138
    vla_step = 663 ' vla:3374
    If vlatraceon() Then ' vla:3375
        Call vlatracestep(663, vla_step_text(663)) ' vla:3375
    End If
    range("j3") = 7 ' vla:3377 src:1139
    vla_step = 664 ' vla:3378
    If vlatraceon() Then ' vla:3379
        Call vlatracestep(664, vla_step_text(664)) ' vla:3379
    End If
    range("j4") = "y" ' vla:3381 src:1140
    vla_step = 665 ' vla:3382
    If vlatraceon() Then ' vla:3383
        Call vlatracestep(665, vla_step_text(665)) ' vla:3383
    End If
    gtext_list = vlajoinrange(range("j1:j4"), ", ") ' vla:3385 src:1141
    vla_step = 666 ' vla:3386
    If vlatraceon() Then ' vla:3387
        Call vlatracestep(666, vla_step_text(666)) ' vla:3387
    End If
    range("i9") = gtext_list ' vla:3389 src:1142
    vla_step = 667 ' vla:3390
    If vlatraceon() Then ' vla:3391
        Call vlatracestep(667, vla_step_text(667)) ' vla:3391
    End If
    gtext_list = vlajoinrange(range("j1:j4"), "; ") ' vla:3393 src:1143
    vla_step = 668 ' vla:3394
    If vlatraceon() Then ' vla:3395
        Call vlatracestep(668, vla_step_text(668)) ' vla:3395
    End If
    range("i10") = gtext_list ' vla:3397 src:1144
    vla_step = 669 ' vla:3398
    If vlatraceon() Then ' vla:3399
        Call vlatracestep(669, vla_step_text(669)) ' vla:3399
    End If
    gtext_list = vlajoinrange(columns("j"), ", ") ' vla:3401 src:1145
    vla_step = 670 ' vla:3402
    If vlatraceon() Then ' vla:3403
        Call vlatracestep(670, vla_step_text(670)) ' vla:3403
    End If
    range("i11") = gtext_list ' vla:3405 src:1146
    vla_step = 671 ' vla:3406
    If vlatraceon() Then ' vla:3407
        Call vlatracestep(671, vla_step_text(671)) ' vla:3407
    End If
    gtext_part = vlatextop("  a   b ", "remove-extra-spaces") ' vla:3409 src:1150
    vla_step = 672 ' vla:3410
    If vlatraceon() Then ' vla:3411
        Call vlatracestep(672, vla_step_text(672)) ' vla:3411
    End If
    range("i12") = gtext_part ' vla:3413 src:1151
    vla_step = 673 ' vla:3414
    If vlatraceon() Then ' vla:3415
        Call vlatracestep(673, vla_step_text(673)) ' vla:3415
    End If
    gtext_part = vlatextop("o'neil", "capitalize-after-non-letter") ' vla:3417 src:1152
    vla_step = 674 ' vla:3418
    If vlatraceon() Then ' vla:3419
        Call vlatracestep(674, vla_step_text(674)) ' vla:3419
    End If
    range("i13") = gtext_part ' vla:3421 src:1153
    vla_step = 675 ' vla:3422
    If vlatraceon() Then ' vla:3423
        Call vlatracestep(675, vla_step_text(675)) ' vla:3423
    End If
    gtext_part = vlatextop("don't stop", "capitalize-after-space") ' vla:3425 src:1154
    vla_step = 676 ' vla:3426
    If vlatraceon() Then ' vla:3427
        Call vlatracestep(676, vla_step_text(676)) ' vla:3427
    End If
    range("i14") = gtext_part ' vla:3429 src:1155
    vla_step = 677 ' vla:3430
    If vlatraceon() Then ' vla:3431
        Call vlatracestep(677, vla_step_text(677)) ' vla:3431
    End If
    range("k1").Formula2 = "=CHAR(112)&CHAR(10)&CHAR(113)"
    vla_step = 678 ' vla:3434
    If vlatraceon() Then ' vla:3435
        Call vlatracestep(678, vla_step_text(678)) ' vla:3435
    End If
    Call range("k1:k1").copy
    Call range("k1:k1").pastespecial(paste:=xlpastevalues)
    application.cutcopymode = False
    vla_step = 679 ' vla:3438
    If vlatraceon() Then ' vla:3439
        Call vlatracestep(679, vla_step_text(679)) ' vla:3439
    End If
    gtext_part = vlatextop(range("k1"), "remove-non-printing") ' vla:3441 src:1158
    vla_step = 680 ' vla:3442
    If vlatraceon() Then ' vla:3443
        Call vlatracestep(680, vla_step_text(680)) ' vla:3443
    End If
    range("i15") = gtext_part ' vla:3445 src:1159
    vla_step = 681 ' vla:3446
    If vlatraceon() Then ' vla:3447
        Call vlatracestep(681, vla_step_text(681)) ' vla:3447
    End If
    On Error GoTo vla_tryf_10 ' vla:3449 src:1164
    vla_step = 682 ' vla:3450 src:1164
    If vlatraceon() Then ' vla:3451 src:1164
        Call vlatracestep(682, vla_step_text(682)) ' vla:3451 src:1164
    End If
    gtext_part = vlatextbeside(gtext_code, "#", "before", "first") ' vla:3453 src:1165
    GoTo vla_tryd_10 ' vla:3454 src:1164
vla_tryf_10: ' vla:3455 src:1164
    vla_problem = err.description ' vla:3456 src:1164
    Resume vla_tryr_10 ' vla:3457 src:1164
vla_tryr_10: ' vla:3458 src:1164
    On Error GoTo vla_fail ' vla:3459 src:1164
vla_tryd_10: ' vla:3460 src:1164
    On Error GoTo vla_fail ' vla:3461 src:1164
    vla_step = 683 ' vla:3462
    If vlatraceon() Then ' vla:3463
        Call vlatracestep(683, vla_step_text(683)) ' vla:3463
    End If
    range("i16") = gtext_part ' vla:3465 src:1167
    vla_step = 684 ' vla:3466
    If vlatraceon() Then ' vla:3467
        Call vlatracestep(684, vla_step_text(684)) ' vla:3467
    End If
    range("l1") = "Nan" ' vla:3469 src:1173
    vla_step = 685 ' vla:3470
    If vlatraceon() Then ' vla:3471
        Call vlatracestep(685, vla_step_text(685)) ' vla:3471
    End If
    range("l2").Formula2 = "=TAN(0)"
    vla_step = 686 ' vla:3474
    If vlatraceon() Then ' vla:3475
        Call vlatracestep(686, vla_step_text(686)) ' vla:3475
    End If
    Call vlareplaceinrange(range("l1:l2"), "an", "xy", "values")
    vla_step = 687 ' vla:3478
    If vlatraceon() Then ' vla:3479
        Call vlatracestep(687, vla_step_text(687)) ' vla:3479
    End If
    range("l3") = "Q" ' vla:3481 src:1176
    vla_step = 688 ' vla:3482
    If vlatraceon() Then ' vla:3483
        Call vlatracestep(688, vla_step_text(688)) ' vla:3483
    End If
    Call vlareplaceinrange(range("l3:l3"), "Q", "=1+1", "values")
    vla_step = 689 ' vla:3486
    If vlatraceon() Then ' vla:3487
        Call vlatracestep(689, vla_step_text(689)) ' vla:3487
    End If
    range("l4") = "N/A" ' vla:3489 src:1178
    vla_step = 690 ' vla:3490
    If vlatraceon() Then ' vla:3491
        Call vlatracestep(690, vla_step_text(690)) ' vla:3491
    End If
    Call vlareplaceinrange(range("l4:l4"), "N/A", 0, "values")
    vla_step = 691 ' vla:3494
    If vlatraceon() Then ' vla:3495
        Call vlatracestep(691, vla_step_text(691)) ' vla:3495
    End If
    range("l5").Formula2 = "=ABS(-2)"
    vla_step = 692 ' vla:3498
    If vlatraceon() Then ' vla:3499
        Call vlatracestep(692, vla_step_text(692)) ' vla:3499
    End If
    range("l6") = "ABS" ' vla:3501 src:1183
    vla_step = 693 ' vla:3502
    If vlatraceon() Then ' vla:3503
        Call vlatracestep(693, vla_step_text(693)) ' vla:3503
    End If
    Call vlareplaceinrange(range("l5:l6"), "ABS", "SIGN", "formulas")
    vla_step = 694 ' vla:3506
    If vlatraceon() Then ' vla:3507
        Call vlatracestep(694, vla_step_text(694)) ' vla:3507
    End If
    range("l7") = "zqz" ' vla:3509 src:1187
    vla_step = 695 ' vla:3510
    If vlatraceon() Then ' vla:3511
        Call vlatracestep(695, vla_step_text(695)) ' vla:3511
    End If
    Call vlareplaceinrange(activesheet.usedrange, "zqz", "done", "values")
    vla_step = 696 ' vla:3514
    If vlatraceon() Then ' vla:3515
        Call vlatracestep(696, vla_step_text(696)) ' vla:3515
    End If
    range("l8").Formula2 = "=ABS(-3)"
    vla_step = 697 ' vla:3518
    If vlatraceon() Then ' vla:3519
        Call vlatracestep(697, vla_step_text(697)) ' vla:3519
    End If
    Call vlareplaceinrange(activesheet.usedrange, "-3", "-4", "formulas")
    vla_step = 698 ' vla:3522
    If vlatraceon() Then ' vla:3523
        Call vlatracestep(698, vla_step_text(698)) ' vla:3523
    End If
    range("n1") = "x,0042,7" ' vla:3525 src:1195
    vla_step = 699 ' vla:3526
    If vlatraceon() Then ' vla:3527
        Call vlatracestep(699, vla_step_text(699)) ' vla:3527
    End If
    Call vlasplitcolumn(columns("n"), ",", "text")
    vla_step = 700 ' vla:3530
    If vlatraceon() Then ' vla:3531
        Call vlatracestep(700, vla_step_text(700)) ' vla:3531
    End If
    range("r1") = "y,0042,7,-3.5" ' vla:3533 src:1197
    vla_step = 701 ' vla:3534
    If vlatraceon() Then ' vla:3535
        Call vlatracestep(701, vla_step_text(701)) ' vla:3535
    End If
    Call vlasplitcolumn(columns("r"), ",", "numbers")
    vla_step = 702 ' vla:3538
    If vlatraceon() Then ' vla:3539
        Call vlatracestep(702, vla_step_text(702)) ' vla:3539
    End If
    range("w1") = "p,q" ' vla:3541 src:1199
    vla_step = 703 ' vla:3542
    If vlatraceon() Then ' vla:3543
        Call vlatracestep(703, vla_step_text(703)) ' vla:3543
    End If
    range("x1") = "keep" ' vla:3545 src:1200
    vla_step = 704 ' vla:3546
    If vlatraceon() Then ' vla:3547
        Call vlatracestep(704, vla_step_text(704)) ' vla:3547
    End If
    On Error GoTo vla_tryf_11 ' vla:3549 src:1201
    vla_step = 705 ' vla:3550 src:1201
    If vlatraceon() Then ' vla:3551 src:1201
        Call vlatracestep(705, vla_step_text(705)) ' vla:3551 src:1201
    End If
    Call vlasplitcolumn(columns("w"), ",", "text")
    GoTo vla_tryd_11 ' vla:3554 src:1201
vla_tryf_11: ' vla:3555 src:1201
    vla_problem = err.description ' vla:3556 src:1201
    Resume vla_tryr_11 ' vla:3557 src:1201
vla_tryr_11: ' vla:3558 src:1201
    On Error GoTo vla_fail ' vla:3559 src:1201
vla_tryd_11: ' vla:3560 src:1201
    On Error GoTo vla_fail ' vla:3561 src:1201
    vla_step = 706 ' vla:3562
    If vlatraceon() Then ' vla:3563
        Call vlatracestep(706, vla_step_text(706)) ' vla:3563
    End If
    range("z1") = "apple" ' vla:3565 src:1205
    vla_step = 707 ' vla:3566
    If vlatraceon() Then ' vla:3567
        Call vlatracestep(707, vla_step_text(707)) ' vla:3567
    End If
    range("aa1") = "Banana" ' vla:3569 src:1206
    vla_step = 708 ' vla:3570
    If vlatraceon() Then ' vla:3571
        Call vlatracestep(708, vla_step_text(708)) ' vla:3571
    End If
    range("z2") = "cherry" ' vla:3573 src:1207
    vla_step = 709 ' vla:3574
    If vlatraceon() Then ' vla:3575
        Call vlatracestep(709, vla_step_text(709)) ' vla:3575
    End If
    range("aa2") = "banana split" ' vla:3577 src:1208
    vla_step = 710 ' vla:3578
    If vlatraceon() Then ' vla:3579
        Call vlatracestep(710, vla_step_text(710)) ' vla:3579
    End If
    range("z3") = 7 ' vla:3581 src:1209
    vla_step = 711 ' vla:3582
    If vlatraceon() Then ' vla:3583
        Call vlatracestep(711, vla_step_text(711)) ' vla:3583
    End If
    range("aa3") = "BANANA" ' vla:3585 src:1210
    vla_step = 712 ' vla:3586
    If vlatraceon() Then ' vla:3587
        Call vlatracestep(712, vla_step_text(712)) ' vla:3587
    End If
    gtext_found = vlafindtext(range("z1:aa3"), "nan", "row") ' vla:3589 src:1211
    vla_step = 713 ' vla:3590
    If vlatraceon() Then ' vla:3591
        Call vlatracestep(713, vla_step_text(713)) ' vla:3591
    End If
    range("l9") = gtext_found ' vla:3593 src:1212
    vla_step = 714 ' vla:3594
    If vlatraceon() Then ' vla:3595
        Call vlatracestep(714, vla_step_text(714)) ' vla:3595
    End If
    gtext_found = vlafindtext(range("z1:aa3"), "nan", "column") ' vla:3597 src:1213
    vla_step = 715 ' vla:3598
    If vlatraceon() Then ' vla:3599
        Call vlatracestep(715, vla_step_text(715)) ' vla:3599
    End If
    range("l10") = gtext_found ' vla:3601 src:1214
    vla_step = 716 ' vla:3602
    If vlatraceon() Then ' vla:3603
        Call vlatracestep(716, vla_step_text(716)) ' vla:3603
    End If
    gtext_found = vlafindtext(columns("z"), "err", "row") ' vla:3605 src:1215
    vla_step = 717 ' vla:3606
    If vlatraceon() Then ' vla:3607
        Call vlatracestep(717, vla_step_text(717)) ' vla:3607
    End If
    range("l11") = gtext_found ' vla:3609 src:1216
    vla_step = 718 ' vla:3610
    If vlatraceon() Then ' vla:3611
        Call vlatracestep(718, vla_step_text(718)) ' vla:3611
    End If
    gtext_found = vlacounttext(range("z1:aa3"), "banana") ' vla:3613 src:1217
    vla_step = 719 ' vla:3614
    If vlatraceon() Then ' vla:3615
        Call vlatracestep(719, vla_step_text(719)) ' vla:3615
    End If
    range("l12") = gtext_found ' vla:3617 src:1218
    vla_step = 720 ' vla:3618
    If vlatraceon() Then ' vla:3619
        Call vlatracestep(720, vla_step_text(720)) ' vla:3619
    End If
    gtext_found = vlacounttext(columns("z"), "zzz") ' vla:3621 src:1219
    vla_step = 721 ' vla:3622
    If vlatraceon() Then ' vla:3623
        Call vlatracestep(721, vla_step_text(721)) ' vla:3623
    End If
    range("l13") = gtext_found ' vla:3625 src:1220
    vla_step = 722 ' vla:3626
    If vlatraceon() Then ' vla:3627
        Call vlatracestep(722, vla_step_text(722)) ' vla:3627
    End If
    gtext_found = vlafindtext(range("z1:aa3"), "kiwi", "row") ' vla:3629 src:1221
    vla_step = 723 ' vla:3630
    If vlatraceon() Then ' vla:3631
        Call vlatracestep(723, vla_step_text(723)) ' vla:3631
    End If
    range("l14") = gtext_found ' vla:3633 src:1222
    vla_step = 724 ' vla:3634
    If vlatraceon() Then ' vla:3635
        Call vlatracestep(724, vla_step_text(724)) ' vla:3635
    End If
    If (vlafindtext(range("z1:aa3"), "split", "row") > 0) Then ' vla:3637 src:1223
        range("l15") = "yes" ' vla:3637 src:1223
    End If
    vla_step = 725 ' vla:3638
    If vlatraceon() Then ' vla:3639
        Call vlatracestep(725, vla_step_text(725)) ' vla:3639
    End If
    If (vlafindtext(columns("z"), "kiwi", "row") = 0) Then ' vla:3641 src:1224
        range("l16") = "no kiwi" ' vla:3641 src:1224
    End If
    vla_step = 735 ' vla:3642
    If vlatraceon() Then ' vla:3643
        Call vlatracestep(735, vla_step_text(735)) ' vla:3643
    End If
    Call check_text_conditions ' vla:3645 src:1248
    vla_step = 761 ' vla:3646
    If vlatraceon() Then ' vla:3647
        Call vlatracestep(761, vla_step_text(761)) ' vla:3647
    End If
    Call check_formulas ' vla:3649 src:1285
    vla_step = 789 ' vla:3650
    If vlatraceon() Then ' vla:3651
        Call vlatracestep(789, vla_step_text(789)) ' vla:3651
    End If
    Call check_statistics ' vla:3653 src:1330
    vla_step = 812 ' vla:3654
    If vlatraceon() Then ' vla:3655
        Call vlatracestep(812, vla_step_text(812)) ' vla:3655
    End If
    Call check_calculation ' vla:3657 src:1366
    vla_step = 842 ' vla:3658
    If vlatraceon() Then ' vla:3659
        Call vlatracestep(842, vla_step_text(842)) ' vla:3659
    End If
    Call check_figures ' vla:3661 src:1410
    vla_step = 858 ' vla:3662
    If vlatraceon() Then ' vla:3663
        Call vlatracestep(858, vla_step_text(858)) ' vla:3663
    End If
    Call check_phrases ' vla:3665 src:1438
    vla_step = 859 ' vla:3666
    If vlatraceon() Then ' vla:3667
        Call vlatracestep(859, vla_step_text(859)) ' vla:3667
    End If
    Call worksheets("output").activate
    vla_step = 860 ' vla:3670
    If vlatraceon() Then ' vla:3671
        Call vlatracestep(860, vla_step_text(860)) ' vla:3671
    End If
    Call tidy_up ' vla:3673 src:1442
    vla_step = 861 ' vla:3674
    If vlatraceon() Then ' vla:3675
        Call vlatracestep(861, vla_step_text(861)) ' vla:3675
    End If
    application.screenupdating = True
    vla_step = 862 ' vla:3678
    If vlatraceon() Then ' vla:3679
        Call vlatracestep(862, vla_step_text(862)) ' vla:3679
    End If
    Debug.Print "report finished" ' vla:3681 src:1444
    Exit Sub ' vla:3682
vla_fail: ' vla:3683
    Call vla_report_error ' vla:3684
End Sub

Public Sub vla_report_error()
    Call vlareportstop(vla_step, vla_step_text(vla_step), err.description) ' vla:3687
End Sub

Public Function vla_step_text(ByVal n As Long) As String
    Select Case n ' vla:3690
        Case 1
            vla_step_text = "Work on sheet Output. [line 18]" ' vla:3691
            Exit Function
        Case 2
            vla_step_text = "Turn off screen updating. [line 20]" ' vla:3692
            Exit Function
        Case 3
            vla_step_text = "Create a number called total. [line 21]" ' vla:3693
            Exit Function
        Case 4
            vla_step_text = "Set total to 0. [line 22]" ' vla:3694
            Exit Function
        Case 5
            vla_step_text = "Put ""Test Report"" into cell A1. [line 23]" ' vla:3695
            Exit Function
        Case 6
            vla_step_text = "Make cell A1 bold. [line 24]" ' vla:3696
            Exit Function
        Case 7
            vla_step_text = "Set font size of cell A1 to 14. [line 25]" ' vla:3697
            Exit Function
        Case 8
            vla_step_text = "Put today into cell D1. [line 26]" ' vla:3698
            Exit Function
        Case 9
            vla_step_text = "Repeat 5 times: [line 28]" ' vla:3699
            Exit Function
        Case 10
            vla_step_text = "Increase total by counter. [line 29]" ' vla:3700
            Exit Function
        Case 11
            vla_step_text = "If counter is divisible by 2, log ""even step "" joined with counter. [line 30]" ' vla:3701
            Exit Function
        Case 12
            vla_step_text = "Put total into cell B2. [line 32]" ' vla:3702
            Exit Function
        Case 13
            vla_step_text = "Put formula ""=B2*2"" into cell B3. [line 33]" ' vla:3703
            Exit Function
        Case 14
            vla_step_text = "Put sum of range ""B2:B3"" into cell B4. [line 37]" ' vla:3704
            Exit Function
        Case 15
            vla_step_text = "Set grand to sum of range B2:B3. [line 38]" ' vla:3705
            Exit Function
        Case 16
            vla_step_text = "Log ""grand is "" joined with grand. [line 39]" ' vla:3706
            Exit Function
        Case 17
            vla_step_text = "If grand is greater than 40: [line 43]" ' vla:3707
            Exit Function
        Case 18
            vla_step_text = "Put ""PASS"" into cell C4. [line 44]" ' vla:3708
            Exit Function
        Case 19
            vla_step_text = "Repeat 2 times: [line 45]" ' vla:3709
            Exit Function
        Case 20
            vla_step_text = "Log ""pass check "" joined with counter. [line 46]" ' vla:3710
            Exit Function
        Case 21
            vla_step_text = "Make cell C4 bold. [line 48]" ' vla:3711
            Exit Function
        Case 22
            vla_step_text = "Put ""CHECK"" into cell C4. [line 51]" ' vla:3712
            Exit Function
        Case 23
            vla_step_text = "If grand is at least 45, make cell C4 yellow. [line 53]" ' vla:3713
            Exit Function
        Case 24
            vla_step_text = "Create a text called label. [line 54]" ' vla:3714
            Exit Function
        Case 25
            vla_step_text = "Set label to ""Total: "" joined with total. [line 55]" ' vla:3715
            Exit Function
        Case 26
            vla_step_text = "Put label into cell A6. [line 56]" ' vla:3716
            Exit Function
        Case 27
            vla_step_text = "Set font-color of cell A6 to hot-pink. [line 57]" ' vla:3717
            Exit Function
        Case 28
            vla_step_text = "Remember range B2:B4 as results. [line 58]" ' vla:3718
            Exit Function
        Case 29
            vla_step_text = "For each r in results, log r. [line 59]" ' vla:3719
            Exit Function
        Case 30
            vla_step_text = "Set biggest to largest of results. [line 60]" ' vla:3720
            Exit Function
        Case 31
            vla_step_text = "Log ""largest result is "" joined with biggest joined with "", label length "" joined with length of label. [line 61]" ' vla:3721
            Exit Function
        Case 32
            vla_step_text = "Repeat 3 times: [line 64]" ' vla:3722
            Exit Function
        Case 33
            vla_step_text = "Put counter times 10 into column E row counter. [line 65]" ' vla:3723
            Exit Function
        Case 34
            vla_step_text = "Set probe to cell in column E row 2. [line 67]" ' vla:3724
            Exit Function
        Case 35
            vla_step_text = "Create a text called num-col-check. [line 68]" ' vla:3725
            Exit Function
        Case 36
            vla_step_text = "If cell in column number 5 row 2 is 20, set num-col-check to ""yes"". [line 69]" ' vla:3726
            Exit Function
        Case 37
            vla_step_text = "Create a text called value-word-check. [line 70]" ' vla:3727
            Exit Function
        Case 38
            vla_step_text = "If value in column number 5 row 3 is 30, set value-word-check to ""ok"". [line 71]" ' vla:3728
            Exit Function
        Case 39
            vla_step_text = "Put ""bang"" into cell Output!H16. [line 72]" ' vla:3729
            Exit Function
        Case 40
            vla_step_text = "Center cell A1. [line 73]" ' vla:3730
            Exit Function
        Case 41
            vla_step_text = "Set width of column D to 24. [line 74]" ' vla:3731
            Exit Function
        Case 42
            vla_step_text = "Set round-check to 3.14159 rounded to 2 decimals. [line 75]" ' vla:3732
            Exit Function
        Case 43
            vla_step_text = "Log ""rounded is "" joined with round-check. [line 76]" ' vla:3733
            Exit Function
        Case 44
            vla_step_text = "Set thousand-check to 1,000 plus 500. [line 77]" ' vla:3734
            Exit Function
        Case 45
            vla_step_text = "Log ""thousands read as "" joined with thousand-check. [line 78]" ' vla:3735
            Exit Function
        Case 46
            vla_step_text = "Log ""probe is "" joined with probe. [line 79]" ' vla:3736
            Exit Function
        Case 47
            vla_step_text = "Create a number called countdown. [line 81]" ' vla:3737
            Exit Function
        Case 48
            vla_step_text = "Set countdown to 3. [line 82]" ' vla:3738
            Exit Function
        Case 49
            vla_step_text = "While countdown is greater than 0: [line 83]" ' vla:3739
            Exit Function
        Case 50
            vla_step_text = "Log ""countdown "" joined with countdown. [line 84]" ' vla:3740
            Exit Function
        Case 51
            vla_step_text = "Decrease countdown by 1. [line 85]" ' vla:3741
            Exit Function
        Case 52
            vla_step_text = "Put value into column F row row-number. [line 90]" ' vla:3742
            Exit Function
        Case 53
            vla_step_text = "Stamp. [line 92]" ' vla:3743
            Exit Function
        Case 54
            vla_step_text = "Stamp with row-number of 2 and value of ""beta"". [line 93]" ' vla:3744
            Exit Function
        Case 55
            vla_step_text = "Set f-last to last filled row of column F. [line 94]" ' vla:3745
            Exit Function
        Case 56
            vla_step_text = "Log ""column F filled to row "" joined with f-last. [line 95]" ' vla:3746
            Exit Function
        Case 57
            vla_step_text = "Count echo-row from 1 to f-last, log ""echo "" joined with echo-row. [line 96]" ' vla:3747
            Exit Function
        Case 58
            vla_step_text = "Count check-row from 1 to f-last: [line 97]" ' vla:3748
            Exit Function
        Case 59
            vla_step_text = "If cell in column F row check-row contains ""ok"", make cell in column F row check-row bold. [line 98]" ' vla:3749
            Exit Function
        Case 60
            vla_step_text = "Count stripe-row from 1 to f-last step 2: [line 103]" ' vla:3750
            Exit Function
        Case 61
            vla_step_text = "Make cell in column F row stripe-row bold. [line 104]" ' vla:3751
            Exit Function
        Case 62
            vla_step_text = "Count back-row down from f-last to 1 step 2: [line 106]" ' vla:3752
            Exit Function
        Case 63
            vla_step_text = "Log ""back-row "" joined with back-row. [line 107]" ' vla:3753
            Exit Function
        Case 64
            vla_step_text = "Count search-row from 1 to 10: [line 111]" ' vla:3754
            Exit Function
        Case 65
            vla_step_text = "If search-row is 3, stop the loop. [line 112]" ' vla:3755
            Exit Function
        Case 66
            vla_step_text = "Log ""stopped at "" joined with search-row. [line 114]" ' vla:3756
            Exit Function
        Case 67
            vla_step_text = "Create a list called found-items. [line 119]" ' vla:3757
            Exit Function
        Case 68
            vla_step_text = "Repeat 4 times: [line 120]" ' vla:3758
            Exit Function
        Case 69
            vla_step_text = "If counter is greater than 2, append counter times 100 to found-items. [line 121]" ' vla:3759
            Exit Function
        Case 70
            vla_step_text = "For each f in found-items, log ""found "" joined with f. [line 123]" ' vla:3760
            Exit Function
        Case 71
            vla_step_text = "Set list-count to count of found-items. [line 124]" ' vla:3761
            Exit Function
        Case 72
            vla_step_text = "Log ""list holds "" joined with list-count. [line 125]" ' vla:3762
            Exit Function
        Case 73
            vla_step_text = "Create a text called verdict. [line 131]" ' vla:3763
            Exit Function
        Case 74
            vla_step_text = "If grand is greater than 100: [line 132]" ' vla:3764
            Exit Function
        Case 75
            vla_step_text = "Set verdict to ""huge"". [line 133]" ' vla:3765
            Exit Function
        Case 76
            vla_step_text = "Set verdict to ""solid"". [line 136]" ' vla:3766
            Exit Function
        Case 77
            vla_step_text = "Set verdict to ""small"". [line 139]" ' vla:3767
            Exit Function
        Case 78
            vla_step_text = "Create a text called region-label. [line 141]" ' vla:3768
            Exit Function
        Case 79
            vla_step_text = "Set region to ""South"". [line 142]" ' vla:3769
            Exit Function
        Case 80
            vla_step_text = "When region is ""North"": [line 143]" ' vla:3770
            Exit Function
        Case 81
            vla_step_text = "Set region-label to ""cold"". [line 144]" ' vla:3771
            Exit Function
        Case 82
            vla_step_text = "Set region-label to ""warm"". [line 147]" ' vla:3772
            Exit Function
        Case 83
            vla_step_text = "Set region-label to ""unknown"". [line 150]" ' vla:3773
            Exit Function
        Case 84
            vla_step_text = "Log ""verdict "" joined with verdict joined with "", region "" joined with region-label. [line 152]" ' vla:3774
            Exit Function
        Case 85
            vla_step_text = "Create a number called until-count. [line 154]" ' vla:3775
            Exit Function
        Case 86
            vla_step_text = "Set fuel to 3. [line 155]" ' vla:3776
            Exit Function
        Case 87
            vla_step_text = "Repeat until fuel is 0: [line 156]" ' vla:3777
            Exit Function
        Case 88
            vla_step_text = "Decrease fuel by 1. [line 157]" ' vla:3778
            Exit Function
        Case 89
            vla_step_text = "Increase until-count by 1. [line 158]" ' vla:3779
            Exit Function
        Case 90
            vla_step_text = "Log ""repeat-until ran "" joined with until-count joined with "" times"". [line 160]" ' vla:3780
            Exit Function
        Case 91
            vla_step_text = "Create a text called rescue. [line 166]" ' vla:3781
            Exit Function
        Case 92
            vla_step_text = "Try: [line 167]" ' vla:3782
            Exit Function
        Case 93
            vla_step_text = "Go to sheet Nowhere-Land. [line 168]" ' vla:3783
            Exit Function
        Case 94
            vla_step_text = "Set rescue to ""unreachable"". [line 169]" ' vla:3784
            Exit Function
        Case 95
            vla_step_text = "Log ""the problem was "" joined with the problem. [line 172]" ' vla:3785
            Exit Function
        Case 96
            vla_step_text = "If the problem is not empty, set rescue to ""rescued"". [line 173]" ' vla:3786
            Exit Function
        Case 97
            vla_step_text = "Create a number called risk-free. [line 175]" ' vla:3787
            Exit Function
        Case 98
            vla_step_text = "Try: [line 176]" ' vla:3788
            Exit Function
        Case 99
            vla_step_text = "Set risk-free to 7. [line 177]" ' vla:3789
            Exit Function
        Case 100
            vla_step_text = "Set risk-free to -1. [line 180]" ' vla:3790
            Exit Function
        Case 101
            vla_step_text = "Log ""rescue "" joined with rescue joined with "", risk-free "" joined with risk-free. [line 182]" ' vla:3791
            Exit Function
        Case 102
            vla_step_text = "Give back amount times 0.08. [line 187]" ' vla:3792
            Exit Function
        Case 103
            vla_step_text = "Set fee to tax of 100. [line 189]" ' vla:3793
            Exit Function
        Case 104
            vla_step_text = "Log ""fee is "" joined with fee. [line 190]" ' vla:3794
            Exit Function
        Case 105
            vla_step_text = "Create a text called fee-size. [line 191]" ' vla:3795
            Exit Function
        Case 106
            vla_step_text = "If tax of 50 is greater than 3: [line 192]" ' vla:3796
            Exit Function
        Case 107
            vla_step_text = "Set fee-size to ""big"". [line 193]" ' vla:3797
            Exit Function
        Case 108
            vla_step_text = "Set fee-size to ""small"". [line 196]" ' vla:3798
            Exit Function
        Case 109
            vla_step_text = "Fit column A. [line 199]" ' vla:3799
            Exit Function
        Case 110
            vla_step_text = "Fit column B. [line 200]" ' vla:3800
            Exit Function
        Case 111
            vla_step_text = "Set font-color of cell A1 to hot-pink. [line 201]" ' vla:3801
            Exit Function
        Case 112
            vla_step_text = "Give back sale times rate. [line 206]" ' vla:3802
            Exit Function
        Case 113
            vla_step_text = "Set full-commission to commission using sale of 2000 and rate of 10%. [line 208]" ' vla:3803
            Exit Function
        Case 114
            vla_step_text = "Set default-commission to get commission using sale of 600. [line 209]" ' vla:3804
            Exit Function
        Case 115
            vla_step_text = "Log ""commissions "" joined with full-commission joined with "" / "" joined with default-commission. [line 210]" ' vla:3805
            Exit Function
        Case 116
            vla_step_text = "Try: [line 219]" ' vla:3806
            Exit Function
        Case 117
            vla_step_text = "Add sheet called ""Q1 Data"". [line 220]" ' vla:3807
            Exit Function
        Case 118
            vla_step_text = "Go to sheet Output. [line 222]" ' vla:3808
            Exit Function
        Case 119
            vla_step_text = "Put ""spaced"" into cell 'Q1 Data'!A1. [line 223]" ' vla:3809
            Exit Function
        Case 120
            vla_step_text = "Create a text called spaced-check. [line 224]" ' vla:3810
            Exit Function
        Case 121
            vla_step_text = "Set spaced-check to value in cell 'Q1 Data'!A1. [line 225]" ' vla:3811
            Exit Function
        Case 122
            vla_step_text = "Give back 20%. [line 230]" ' vla:3812
            Exit Function
        Case 123
            vla_step_text = "Create a number called growth-check. [line 237]" ' vla:3813
            Exit Function
        Case 124
            vla_step_text = "Set growth-check to 200. [line 238]" ' vla:3814
            Exit Function
        Case 125
            vla_step_text = "Grow growth-check by 10%. [line 239]" ' vla:3815
            Exit Function
        Case 126
            vla_step_text = "Increase growth-check by 50%. [line 240]" ' vla:3816
            Exit Function
        Case 127
            vla_step_text = "Add 100 percent to growth-check. [line 241]" ' vla:3817
            Exit Function
        Case 128
            vla_step_text = "Decrease growth-check by 75%. [line 242]" ' vla:3818
            Exit Function
        Case 129
            vla_step_text = "Set pick-check to item 2 of found-items. [line 243]" ' vla:3819
            Exit Function
        Case 130
            vla_step_text = "Log ""grew to "" joined with growth-check joined with "", picked "" joined with pick-check joined with "", first "" joined with first of found-items. [line 244]" ' vla:3820
            Exit Function
        Case 131
            vla_step_text = "Create a text called quote-check. [line 245]" ' vla:3821
            Exit Function
        Case 132
            vla_step_text = "Set quote-check to ""He said """"ok"""""". [line 246]" ' vla:3822
            Exit Function
        Case 133
            vla_step_text = "Create a lookup called prices. [line 260]" ' vla:3823
            Exit Function
        Case 134
            vla_step_text = "Store 100 at key ""ax-7"" in prices. [line 261]" ' vla:3824
            Exit Function
        Case 135
            vla_step_text = "Store 250 under ""bx-2"" in prices. [line 262]" ' vla:3825
            Exit Function
        Case 136
            vla_step_text = "Store 120 at key ""AX-7"" in prices. [line 263]" ' vla:3826
            Exit Function
        Case 137
            vla_step_text = "Set ax-price to prices for ""ax-7"". [line 264]" ' vla:3827
            Exit Function
        Case 138
            vla_step_text = "Set key-count to count of keys of prices. [line 265]" ' vla:3828
            Exit Function
        Case 139
            vla_step_text = "Set price-sum to prices for ""ax-7"" plus prices for ""bx-2"". [line 266]" ' vla:3829
            Exit Function
        Case 140
            vla_step_text = "Create a text called key-list. [line 267]" ' vla:3830
            Exit Function
        Case 141
            vla_step_text = "For each k in keys of prices: [line 268]" ' vla:3831
            Exit Function
        Case 142
            vla_step_text = "Set key-list to key-list joined with k. [line 269]" ' vla:3832
            Exit Function
        Case 143
            vla_step_text = "Create a text called price-verdict. [line 271]" ' vla:3833
            Exit Function
        Case 144
            vla_step_text = "If prices for ""bx-2"" is greater than 200, set price-verdict to ""steep"". [line 272]" ' vla:3834
            Exit Function
        Case 145
            vla_step_text = "Create a text called pair-trace. [line 273]" ' vla:3835
            Exit Function
        Case 146
            vla_step_text = "For each pair in prices: [line 274]" ' vla:3836
            Exit Function
        Case 147
            vla_step_text = "If value of pair is greater than 200, set pair-trace to pair-trace joined with key of pair. [line 275]" ' vla:3837
            Exit Function
        Case 148
            vla_step_text = "Log ""lookup: ax "" joined with ax-price joined with "", keys "" joined with key-list joined with "", pairs "" joined with pair-trace. [line 277]" ' vla:3838
            Exit Function
        Case 149
            vla_step_text = "Put ""Item"" into cell J1. [line 284]" ' vla:3839
            Exit Function
        Case 150
            vla_step_text = "Put ""Amount"" into cell K1. [line 285]" ' vla:3840
            Exit Function
        Case 151
            vla_step_text = "Put ""Widget"" into cell J2. [line 286]" ' vla:3841
            Exit Function
        Case 152
            vla_step_text = "Put 10 into cell K2. [line 287]" ' vla:3842
            Exit Function
        Case 153
            vla_step_text = "Put ""Gadget"" into cell J3. [line 288]" ' vla:3843
            Exit Function
        Case 154
            vla_step_text = "Put 20 into cell K3. [line 289]" ' vla:3844
            Exit Function
        Case 155
            vla_step_text = "Turn J1:K3 into a table called SalesTable. [line 290]" ' vla:3845
            Exit Function
        Case 156
            vla_step_text = "Set style of table SalesTable to TableStyleMedium9. [line 291]" ' vla:3846
            Exit Function
        Case 157
            vla_step_text = "Show the total row of table SalesTable. [line 292]" ' vla:3847
            Exit Function
        Case 158
            vla_step_text = "Put ""X"" into cell J5. [line 296]" ' vla:3848
            Exit Function
        Case 159
            vla_step_text = "Put ""Y"" into cell J6. [line 297]" ' vla:3849
            Exit Function
        Case 160
            vla_step_text = "Turn J5:J6 into a table called QuietTable. [line 298]" ' vla:3850
            Exit Function
        Case 161
            vla_step_text = "Show the total row of table QuietTable. [line 299]" ' vla:3851
            Exit Function
        Case 162
            vla_step_text = "Hide the total row of table QuietTable. [line 300]" ' vla:3852
            Exit Function
        Case 163
            vla_step_text = "Put ""A"" into cell J8. [line 304]" ' vla:3853
            Exit Function
        Case 164
            vla_step_text = "Put ""B"" into cell J9. [line 305]" ' vla:3854
            Exit Function
        Case 165
            vla_step_text = "Turn J8:J9 into a table called TempTable. [line 306]" ' vla:3855
            Exit Function
        Case 166
            vla_step_text = "Turn table TempTable back into a range. [line 307]" ' vla:3856
            Exit Function
        Case 167
            vla_step_text = "Put ""Item"" into cell J11. [line 312]" ' vla:3857
            Exit Function
        Case 168
            vla_step_text = "Put ""Qty"" into cell K11. [line 313]" ' vla:3858
            Exit Function
        Case 169
            vla_step_text = "Put ""Bolt"" into cell J12. [line 314]" ' vla:3859
            Exit Function
        Case 170
            vla_step_text = "Put 5 into cell K12. [line 315]" ' vla:3860
            Exit Function
        Case 171
            vla_step_text = "Put ""Nut"" into cell J13. [line 316]" ' vla:3861
            Exit Function
        Case 172
            vla_step_text = "Put 8 into cell K13. [line 317]" ' vla:3862
            Exit Function
        Case 173
            vla_step_text = "Turn J11:K13 into a table called EditTable. [line 318]" ' vla:3863
            Exit Function
        Case 174
            vla_step_text = "Add a row to table EditTable. [line 319]" ' vla:3864
            Exit Function
        Case 175
            vla_step_text = "Delete row 1 of table EditTable. [line 320]" ' vla:3865
            Exit Function
        Case 176
            vla_step_text = "Set qty-values to column Qty of table EditTable. [line 321]" ' vla:3866
            Exit Function
        Case 177
            vla_step_text = "For each q in qty-values, log q. [line 322]" ' vla:3867
            Exit Function
        Case 178
            vla_step_text = "Put ""Region"" into cell N1. [line 332]" ' vla:3868
            Exit Function
        Case 179
            vla_step_text = "Put ""Product"" into cell O1. [line 333]" ' vla:3869
            Exit Function
        Case 180
            vla_step_text = "Put ""Segment"" into cell P1. [line 334]" ' vla:3870
            Exit Function
        Case 181
            vla_step_text = "Put ""Channel"" into cell Q1. [line 335]" ' vla:3871
            Exit Function
        Case 182
            vla_step_text = "Put ""Units"" into cell R1. [line 336]" ' vla:3872
            Exit Function
        Case 183
            vla_step_text = "Put ""Revenue"" into cell S1. [line 337]" ' vla:3873
            Exit Function
        Case 184
            vla_step_text = "Put ""North"" into cell N2. [line 338]" ' vla:3874
            Exit Function
        Case 185
            vla_step_text = "Put ""Widget"" into cell O2. [line 339]" ' vla:3875
            Exit Function
        Case 186
            vla_step_text = "Put ""Retail"" into cell P2. [line 340]" ' vla:3876
            Exit Function
        Case 187
            vla_step_text = "Put ""Online"" into cell Q2. [line 341]" ' vla:3877
            Exit Function
        Case 188
            vla_step_text = "Put 10 into cell R2. [line 342]" ' vla:3878
            Exit Function
        Case 189
            vla_step_text = "Put 500 into cell S2. [line 343]" ' vla:3879
            Exit Function
        Case 190
            vla_step_text = "Put ""North"" into cell N3. [line 344]" ' vla:3880
            Exit Function
        Case 191
            vla_step_text = "Put ""Gadget"" into cell O3. [line 345]" ' vla:3881
            Exit Function
        Case 192
            vla_step_text = "Put ""Wholesale"" into cell P3. [line 346]" ' vla:3882
            Exit Function
        Case 193
            vla_step_text = "Put ""Store"" into cell Q3. [line 347]" ' vla:3883
            Exit Function
        Case 194
            vla_step_text = "Put 5 into cell R3. [line 348]" ' vla:3884
            Exit Function
        Case 195
            vla_step_text = "Put 200 into cell S3. [line 349]" ' vla:3885
            Exit Function
        Case 196
            vla_step_text = "Put ""South"" into cell N4. [line 350]" ' vla:3886
            Exit Function
        Case 197
            vla_step_text = "Put ""Widget"" into cell O4. [line 351]" ' vla:3887
            Exit Function
        Case 198
            vla_step_text = "Put ""Wholesale"" into cell P4. [line 352]" ' vla:3888
            Exit Function
        Case 199
            vla_step_text = "Put ""Online"" into cell Q4. [line 353]" ' vla:3889
            Exit Function
        Case 200
            vla_step_text = "Put 20 into cell R4. [line 354]" ' vla:3890
            Exit Function
        Case 201
            vla_step_text = "Put 900 into cell S4. [line 355]" ' vla:3891
            Exit Function
        Case 202
            vla_step_text = "Put ""South"" into cell N5. [line 356]" ' vla:3892
            Exit Function
        Case 203
            vla_step_text = "Put ""Gadget"" into cell O5. [line 357]" ' vla:3893
            Exit Function
        Case 204
            vla_step_text = "Put ""Retail"" into cell P5. [line 358]" ' vla:3894
            Exit Function
        Case 205
            vla_step_text = "Put ""Store"" into cell Q5. [line 359]" ' vla:3895
            Exit Function
        Case 206
            vla_step_text = "Put 8 into cell R5. [line 360]" ' vla:3896
            Exit Function
        Case 207
            vla_step_text = "Put 300 into cell S5. [line 361]" ' vla:3897
            Exit Function
        Case 208
            vla_step_text = "Put ""East"" into cell N6. [line 362]" ' vla:3898
            Exit Function
        Case 209
            vla_step_text = "Put ""Widget"" into cell O6. [line 363]" ' vla:3899
            Exit Function
        Case 210
            vla_step_text = "Put ""Retail"" into cell P6. [line 364]" ' vla:3900
            Exit Function
        Case 211
            vla_step_text = "Put ""Online"" into cell Q6. [line 365]" ' vla:3901
            Exit Function
        Case 212
            vla_step_text = "Put 12 into cell R6. [line 366]" ' vla:3902
            Exit Function
        Case 213
            vla_step_text = "Put 600 into cell S6. [line 367]" ' vla:3903
            Exit Function
        Case 214
            vla_step_text = "Put ""East"" into cell N7. [line 368]" ' vla:3904
            Exit Function
        Case 215
            vla_step_text = "Put ""Gadget"" into cell O7. [line 369]" ' vla:3905
            Exit Function
        Case 216
            vla_step_text = "Put ""Wholesale"" into cell P7. [line 370]" ' vla:3906
            Exit Function
        Case 217
            vla_step_text = "Put ""Store"" into cell Q7. [line 371]" ' vla:3907
            Exit Function
        Case 218
            vla_step_text = "Put 6 into cell R7. [line 372]" ' vla:3908
            Exit Function
        Case 219
            vla_step_text = "Put 250 into cell S7. [line 373]" ' vla:3909
            Exit Function
        Case 220
            vla_step_text = "Make a pivot table from N1:S7 at U1 called SalesPivot. [line 381]" ' vla:3910
            Exit Function
        Case 221
            vla_step_text = "Add rows of Region, Product to pivot SalesPivot. [line 382]" ' vla:3911
            Exit Function
        Case 222
            vla_step_text = "Add columns of Segment to pivot SalesPivot. [line 383]" ' vla:3912
            Exit Function
        Case 223
            vla_step_text = "Add filters of Channel to pivot SalesPivot. [line 384]" ' vla:3913
            Exit Function
        Case 224
            vla_step_text = "Add Revenue to pivot SalesPivot as a sum. [line 392]" ' vla:3914
            Exit Function
        Case 225
            vla_step_text = "Add Revenue, Units to pivot SalesPivot as a count. [line 393]" ' vla:3915
            Exit Function
        Case 226
            vla_step_text = "Add Units to pivot SalesPivot as an average. [line 394]" ' vla:3916
            Exit Function
        Case 227
            vla_step_text = "Make a pivot table from N1:S7 at N20 called FullPivot with rows of Region, Product, and Channel and columns of Segment and values of Revenue, Units. [line 406]" ' vla:3917
            Exit Function
        Case 228
            vla_step_text = "Refresh pivot SalesPivot. [line 415]" ' vla:3918
            Exit Function
        Case 229
            vla_step_text = "Refresh every pivot table. [line 416]" ' vla:3919
            Exit Function
        Case 230
            vla_step_text = "Collapse Region in pivot SalesPivot. [line 417]" ' vla:3920
            Exit Function
        Case 231
            vla_step_text = "Collapse Product in pivot FullPivot. [line 427]" ' vla:3921
            Exit Function
        Case 232
            vla_step_text = "Expand Product in pivot FullPivot. [line 428]" ' vla:3922
            Exit Function
        Case 233
            vla_step_text = "Show pivot SalesPivot in tabular form. [line 443]" ' vla:3923
            Exit Function
        Case 234
            vla_step_text = "Show pivot SalesPivot in compact form. [line 444]" ' vla:3924
            Exit Function
        Case 235
            vla_step_text = "Show pivot FullPivot in outline form. [line 445]" ' vla:3925
            Exit Function
        Case 236
            vla_step_text = "Hide subtotals for Region in pivot SalesPivot. [line 453]" ' vla:3926
            Exit Function
        Case 237
            vla_step_text = "Hide subtotals for Product in pivot FullPivot. [line 454]" ' vla:3927
            Exit Function
        Case 238
            vla_step_text = "Show subtotals for Product in pivot FullPivot. [line 455]" ' vla:3928
            Exit Function
        Case 239
            vla_step_text = "Add a blank row after Region in pivot SalesPivot. [line 457]" ' vla:3929
            Exit Function
        Case 240
            vla_step_text = "Add a blank row after Product in pivot FullPivot. [line 458]" ' vla:3930
            Exit Function
        Case 241
            vla_step_text = "Remove a blank row after Product in pivot FullPivot. [line 459]" ' vla:3931
            Exit Function
        Case 242
            vla_step_text = "Sort Region in pivot SalesPivot descending. [line 476]" ' vla:3932
            Exit Function
        Case 243
            vla_step_text = "Sort Region in pivot SalesPivot ascending. [line 477]" ' vla:3933
            Exit Function
        Case 244
            vla_step_text = "Sort Product in pivot FullPivot descending by Revenue. [line 478]" ' vla:3934
            Exit Function
        Case 245
            vla_step_text = "Sort Product in pivot FullPivot ascending by Revenue. [line 479]" ' vla:3935
            Exit Function
        Case 246
            vla_step_text = "Make a pivot table from N1:S7 at N200 called RenameMePivot. [line 490]" ' vla:3936
            Exit Function
        Case 247
            vla_step_text = "Rename pivot RenameMePivot to RenamedPivot. [line 491]" ' vla:3937
            Exit Function
        Case 248
            vla_step_text = "Make a pivot table from N1:S7 at N220 called ClearMePivot. [line 493]" ' vla:3938
            Exit Function
        Case 249
            vla_step_text = "Add rows of Region to pivot ClearMePivot. [line 494]" ' vla:3939
            Exit Function
        Case 250
            vla_step_text = "Add Revenue to pivot ClearMePivot as a sum. [line 495]" ' vla:3940
            Exit Function
        Case 251
            vla_step_text = "Clear pivot ClearMePivot. [line 496]" ' vla:3941
            Exit Function
        Case 252
            vla_step_text = "Make a pivot table from N1:S7 at N240 called RemoveFieldMePivot. [line 498]" ' vla:3942
            Exit Function
        Case 253
            vla_step_text = "Add rows of Region, Product to pivot RemoveFieldMePivot. [line 499]" ' vla:3943
            Exit Function
        Case 254
            vla_step_text = "Remove Product from pivot RemoveFieldMePivot. [line 500]" ' vla:3944
            Exit Function
        Case 255
            vla_step_text = "Put ""West"" into cell N8. [line 506]" ' vla:3945
            Exit Function
        Case 256
            vla_step_text = "Put ""Widget"" into cell O8. [line 507]" ' vla:3946
            Exit Function
        Case 257
            vla_step_text = "Put ""Retail"" into cell P8. [line 508]" ' vla:3947
            Exit Function
        Case 258
            vla_step_text = "Put ""Online"" into cell Q8. [line 509]" ' vla:3948
            Exit Function
        Case 259
            vla_step_text = "Put 15 into cell R8. [line 510]" ' vla:3949
            Exit Function
        Case 260
            vla_step_text = "Put 700 into cell S8. [line 511]" ' vla:3950
            Exit Function
        Case 261
            vla_step_text = "Make a pivot table from N1:S7 at N260 called SourceTestPivot. [line 513]" ' vla:3951
            Exit Function
        Case 262
            vla_step_text = "Add rows of Region to pivot SourceTestPivot. [line 514]" ' vla:3952
            Exit Function
        Case 263
            vla_step_text = "Change the source of pivot SourceTestPivot to N1:S8. [line 515]" ' vla:3953
            Exit Function
        Case 264
            vla_step_text = "Make a pivot table from N1:S7 at N40 called TempPivot. [line 520]" ' vla:3954
            Exit Function
        Case 265
            vla_step_text = "Delete pivot TempPivot. [line 521]" ' vla:3955
            Exit Function
        Case 266
            vla_step_text = "Put grand into cell H1. [line 524]" ' vla:3956
            Exit Function
        Case 267
            vla_step_text = "Put biggest into cell H2. [line 525]" ' vla:3957
            Exit Function
        Case 268
            vla_step_text = "Put round-check into cell H3. [line 526]" ' vla:3958
            Exit Function
        Case 269
            vla_step_text = "Put thousand-check into cell H4. [line 527]" ' vla:3959
            Exit Function
        Case 270
            vla_step_text = "Put search-row into cell H5. [line 528]" ' vla:3960
            Exit Function
        Case 271
            vla_step_text = "Put f-last into cell H6. [line 529]" ' vla:3961
            Exit Function
        Case 272
            vla_step_text = "Put list-count into cell H7. [line 530]" ' vla:3962
            Exit Function
        Case 273
            vla_step_text = "Put verdict into cell H8. [line 531]" ' vla:3963
            Exit Function
        Case 274
            vla_step_text = "Put region-label into cell H9. [line 532]" ' vla:3964
            Exit Function
        Case 275
            vla_step_text = "Put until-count into cell H10. [line 533]" ' vla:3965
            Exit Function
        Case 276
            vla_step_text = "Put rescue into cell H11. [line 534]" ' vla:3966
            Exit Function
        Case 277
            vla_step_text = "Put risk-free into cell H12. [line 535]" ' vla:3967
            Exit Function
        Case 278
            vla_step_text = "Put fee into cell H13. [line 536]" ' vla:3968
            Exit Function
        Case 279
            vla_step_text = "Put fee-size into cell H14. [line 537]" ' vla:3969
            Exit Function
        Case 280
            vla_step_text = "Put num-col-check into cell H15. [line 538]" ' vla:3970
            Exit Function
        Case 281
            vla_step_text = "Put value-word-check into cell H17. [line 539]" ' vla:3971
            Exit Function
        Case 282
            vla_step_text = "Put spaced-check into cell H18. [line 540]" ' vla:3972
            Exit Function
        Case 283
            vla_step_text = "Put full-commission into cell H19. [line 541]" ' vla:3973
            Exit Function
        Case 284
            vla_step_text = "Put default-commission into cell H20. [line 542]" ' vla:3974
            Exit Function
        Case 285
            vla_step_text = "Put growth-check into cell H21. [line 543]" ' vla:3975
            Exit Function
        Case 286
            vla_step_text = "Put pick-check into cell H22. [line 544]" ' vla:3976
            Exit Function
        Case 287
            vla_step_text = "Put quote-check into cell H23. [line 545]" ' vla:3977
            Exit Function
        Case 288
            vla_step_text = "Put vat-rate into cell H24. [line 546]" ' vla:3978
            Exit Function
        Case 289
            vla_step_text = "Put ax-price into cell H26. [line 547]" ' vla:3979
            Exit Function
        Case 290
            vla_step_text = "Put key-count into cell H27. [line 548]" ' vla:3980
            Exit Function
        Case 291
            vla_step_text = "Put key-list into cell H28. [line 549]" ' vla:3981
            Exit Function
        Case 292
            vla_step_text = "Put price-sum into cell H29. [line 550]" ' vla:3982
            Exit Function
        Case 293
            vla_step_text = "Put price-verdict into cell H30. [line 551]" ' vla:3983
            Exit Function
        Case 294
            vla_step_text = "Put pair-trace into cell H31. [line 552]" ' vla:3984
            Exit Function
        Case 295
            vla_step_text = "(set! (range ""h25"") ""vla-row"") [line 559]" ' vla:3985
            Exit Function
        Case 296
            vla_step_text = "Repeat 2 times: [line 560]" ' vla:3986
            Exit Function
        Case 297
            vla_step_text = "(debug-print counter) [line 561]" ' vla:3987
            Exit Function
        Case 298
            vla_step_text = "Work on sheet Demo. [line 576]" ' vla:3988
            Exit Function
        Case 299
            vla_step_text = "Make cell A1 italic. [line 579]" ' vla:3989
            Exit Function
        Case 300
            vla_step_text = "Make cell A2 red. [line 580]" ' vla:3990
            Exit Function
        Case 301
            vla_step_text = "Make cell A3 yellow. [line 581]" ' vla:3991
            Exit Function
        Case 302
            vla_step_text = "Set font-color of cell A4 to hot-pink. [line 582]" ' vla:3992
            Exit Function
        Case 303
            vla_step_text = "Set fill-color of cell A5 to ""#FF69B4"". [line 583]" ' vla:3993
            Exit Function
        Case 304
            vla_step_text = "Clear color of cell A5. [line 584]" ' vla:3994
            Exit Function
        Case 305
            vla_step_text = "Add border to range A1:E10. [line 585]" ' vla:3995
            Exit Function
        Case 306
            vla_step_text = "Format cell B1 as currency. [line 586]" ' vla:3996
            Exit Function
        Case 307
            vla_step_text = "Format cell B2 as percent. [line 587]" ' vla:3997
            Exit Function
        Case 308
            vla_step_text = "Format cell B3 as date. [line 588]" ' vla:3998
            Exit Function
        Case 309
            vla_step_text = "Center cell A1. [line 589]" ' vla:3999
            Exit Function
        Case 310
            vla_step_text = "Center range A1:E1. [line 590]" ' vla:4000
            Exit Function
        Case 311
            vla_step_text = "Align range A2:A5 left. [line 591]" ' vla:4001
            Exit Function
        Case 312
            vla_step_text = "Align cell A6 center. [line 592]" ' vla:4002
            Exit Function
        Case 313
            vla_step_text = "Wrap text in range C1:C5. [line 593]" ' vla:4003
            Exit Function
        Case 314
            vla_step_text = "Unwrap text in range C1:C5. [line 594]" ' vla:4004
            Exit Function
        Case 315
            vla_step_text = "Merge range D1:D3. [line 595]" ' vla:4005
            Exit Function
        Case 316
            vla_step_text = "Unmerge range D1:D3. [line 596]" ' vla:4006
            Exit Function
        Case 317
            vla_step_text = "Set height of row 1 to 30. [line 601]" ' vla:4007
            Exit Function
        Case 318
            vla_step_text = "Set width of column A to 20. [line 602]" ' vla:4008
            Exit Function
        Case 319
            vla_step_text = "Insert column before B. [line 603]" ' vla:4009
            Exit Function
        Case 320
            vla_step_text = "Hide column C. [line 604]" ' vla:4010
            Exit Function
        Case 321
            vla_step_text = "Unhide column C. [line 605]" ' vla:4011
            Exit Function
        Case 322
            vla_step_text = "Delete column B. [line 606]" ' vla:4012
            Exit Function
        Case 323
            vla_step_text = "Insert row at 5. [line 607]" ' vla:4013
            Exit Function
        Case 324
            vla_step_text = "Insert a row. [line 608]" ' vla:4014
            Exit Function
        Case 325
            vla_step_text = "Delete row 2. [line 609]" ' vla:4015
            Exit Function
        Case 326
            vla_step_text = "Delete the third row. [line 610]" ' vla:4016
            Exit Function
        Case 327
            vla_step_text = "Hide row 10. [line 611]" ' vla:4017
            Exit Function
        Case 328
            vla_step_text = "Unhide row 10. [line 612]" ' vla:4018
            Exit Function
        Case 329
            vla_step_text = "Freeze top row. [line 613]" ' vla:4019
            Exit Function
        Case 330
            vla_step_text = "Unfreeze panes. [line 614]" ' vla:4020
            Exit Function
        Case 331
            vla_step_text = "Fit column A. [line 615]" ' vla:4021
            Exit Function
        Case 332
            vla_step_text = "Fit all columns. [line 616]" ' vla:4022
            Exit Function
        Case 333
            vla_step_text = "Put ""Region"" into cell G1. [line 620]" ' vla:4023
            Exit Function
        Case 334
            vla_step_text = "Put ""Amount"" into cell H1. [line 621]" ' vla:4024
            Exit Function
        Case 335
            vla_step_text = "Put ""Notes"" into cell I1. [line 622]" ' vla:4025
            Exit Function
        Case 336
            vla_step_text = "Put ""West"" into cell G2. [line 623]" ' vla:4026
            Exit Function
        Case 337
            vla_step_text = "Put 100 into cell H2. [line 624]" ' vla:4027
            Exit Function
        Case 338
            vla_step_text = "Put ""ok"" into cell I2. [line 625]" ' vla:4028
            Exit Function
        Case 339
            vla_step_text = "Put ""East"" into cell G3. [line 626]" ' vla:4029
            Exit Function
        Case 340
            vla_step_text = "Put 250 into cell H3. [line 627]" ' vla:4030
            Exit Function
        Case 341
            vla_step_text = "Put ""ok"" into cell I3. [line 628]" ' vla:4031
            Exit Function
        Case 342
            vla_step_text = "Put ""West"" into cell G4. [line 629]" ' vla:4032
            Exit Function
        Case 343
            vla_step_text = "Put 100 into cell H4. [line 630]" ' vla:4033
            Exit Function
        Case 344
            vla_step_text = "Put ""dup"" into cell I4. [line 631]" ' vla:4034
            Exit Function
        Case 345
            vla_step_text = "Sort range G1:I4 by column H1 descending. [line 632]" ' vla:4035
            Exit Function
        Case 346
            vla_step_text = "Keep only rows of range G1:I4 where column 1 is ""West"". [line 633]" ' vla:4036
            Exit Function
        Case 347
            vla_step_text = "Show all rows. [line 634]" ' vla:4037
            Exit Function
        Case 348
            vla_step_text = "Replace ""dup"" with ""ok"" in range G1:I4. [line 635]" ' vla:4038
            Exit Function
        Case 349
            vla_step_text = "Remove duplicates from range G1:I4 by column 1. [line 636]" ' vla:4039
            Exit Function
        Case 350
            vla_step_text = "Replace ""ok"" with ""fine"" in column I. [line 637]" ' vla:4040
            Exit Function
        Case 351
            vla_step_text = "Convert range G1:I4 to values. [line 638]" ' vla:4041
            Exit Function
        Case 352
            vla_step_text = "Clear formatting of range G1:I4. [line 639]" ' vla:4042
            Exit Function
        Case 353
            vla_step_text = "Name range G1:I4 as demo_table. [line 640]" ' vla:4043
            Exit Function
        Case 354
            vla_step_text = "Copy range G1:I4 to range K1:M4. [line 641]" ' vla:4044
            Exit Function
        Case 355
            vla_step_text = "Set tab-color of sheet Demo to hot-pink. [line 645]" ' vla:4045
            Exit Function
        Case 356
            vla_step_text = "Protect this sheet with password ""demo123"". [line 646]" ' vla:4046
            Exit Function
        Case 357
            vla_step_text = "Unprotect this sheet with password ""demo123"". [line 647]" ' vla:4047
            Exit Function
        Case 358
            vla_step_text = "Try: [line 675]" ' vla:4048
            Exit Function
        Case 359
            vla_step_text = "Delete sheet GStruct. [line 676]" ' vla:4049
            Exit Function
        Case 360
            vla_step_text = "Work on sheet GStruct. [line 678]" ' vla:4050
            Exit Function
        Case 361
            vla_step_text = "Hide row 3. [line 680]" ' vla:4051
            Exit Function
        Case 362
            vla_step_text = "Hide column B. [line 681]" ' vla:4052
            Exit Function
        Case 363
            vla_step_text = "Unhide all rows and columns. [line 682]" ' vla:4053
            Exit Function
        Case 364
            vla_step_text = "Set font size of cell A6 to 36. [line 684]" ' vla:4054
            Exit Function
        Case 365
            vla_step_text = "Fit row 6. [line 685]" ' vla:4055
            Exit Function
        Case 366
            vla_step_text = "Group rows 10 through 12. [line 687]" ' vla:4056
            Exit Function
        Case 367
            vla_step_text = "Group rows 14 through 16. [line 688]" ' vla:4057
            Exit Function
        Case 368
            vla_step_text = "Ungroup rows 14 through 16. [line 689]" ' vla:4058
            Exit Function
        Case 369
            vla_step_text = "Put ""before-insert"" into cell A20. [line 691]" ' vla:4059
            Exit Function
        Case 370
            vla_step_text = "Insert 3 rows at row 20. [line 692]" ' vla:4060
            Exit Function
        Case 371
            vla_step_text = "Put ""before-delete"" into cell A40. [line 694]" ' vla:4061
            Exit Function
        Case 372
            vla_step_text = "Delete rows 38 through 39. [line 695]" ' vla:4062
            Exit Function
        Case 373
            vla_step_text = "Put ""marker-b"" into cell B50. [line 697]" ' vla:4063
            Exit Function
        Case 374
            vla_step_text = "Put ""marker-c"" into cell C50. [line 698]" ' vla:4064
            Exit Function
        Case 375
            vla_step_text = "Put ""marker-d"" into cell D50. [line 699]" ' vla:4065
            Exit Function
        Case 376
            vla_step_text = "Move column B before column D. [line 700]" ' vla:4066
            Exit Function
        Case 377
            vla_step_text = "Freeze the first 2 rows. [line 702]" ' vla:4067
            Exit Function
        Case 378
            vla_step_text = "Make cell A60 bold. [line 705]" ' vla:4068
            Exit Function
        Case 379
            vla_step_text = "Make cell A60 red. [line 706]" ' vla:4069
            Exit Function
        Case 380
            vla_step_text = "Copy formatting of A60:A60 to C60:C60. [line 707]" ' vla:4070
            Exit Function
        Case 381
            vla_step_text = "Make C62:C62 look like A60:A60. [line 708]" ' vla:4071
            Exit Function
        Case 382
            vla_step_text = "Put formula ""=5*2"" into cell B64. [line 710]" ' vla:4072
            Exit Function
        Case 383
            vla_step_text = "Copy formulas of B64 to D64. [line 711]" ' vla:4073
            Exit Function
        Case 384
            vla_step_text = "Set width of column F to 33. [line 713]" ' vla:4074
            Exit Function
        Case 385
            vla_step_text = "Copy column widths of F1:F1 to H1:H1. [line 714]" ' vla:4075
            Exit Function
        Case 386
            vla_step_text = "Put 1 into cell A68. [line 716]" ' vla:4076
            Exit Function
        Case 387
            vla_step_text = "Put 2 into cell A69. [line 717]" ' vla:4077
            Exit Function
        Case 388
            vla_step_text = "Put 3 into cell A70. [line 718]" ' vla:4078
            Exit Function
        Case 389
            vla_step_text = "Copy range A68:A70 to C68 transposed. [line 719]" ' vla:4079
            Exit Function
        Case 390
            vla_step_text = "Put ""cutme"" into cell A72. [line 721]" ' vla:4080
            Exit Function
        Case 391
            vla_step_text = "Move range A72:A72 to C72. [line 722]" ' vla:4081
            Exit Function
        Case 392
            vla_step_text = "Put ""rowdata"" into cell A74. [line 724]" ' vla:4082
            Exit Function
        Case 393
            vla_step_text = "Copy row 74 to row 76. [line 725]" ' vla:4083
            Exit Function
        Case 394
            vla_step_text = "Put ""clearme"" into cell A80. [line 728]" ' vla:4084
            Exit Function
        Case 395
            vla_step_text = "Make cell A80 bold. [line 729]" ' vla:4085
            Exit Function
        Case 396
            vla_step_text = "Make cell A80 red. [line 730]" ' vla:4086
            Exit Function
        Case 397
            vla_step_text = "Clear everything from A80:B80. [line 731]" ' vla:4087
            Exit Function
        Case 398
            vla_step_text = "Put ""m1"" into cell A84. [line 733]" ' vla:4088
            Exit Function
        Case 399
            vla_step_text = "Put ""m2"" into cell A85. [line 734]" ' vla:4089
            Exit Function
        Case 400
            vla_step_text = "Put ""m3"" into cell A86. [line 735]" ' vla:4090
            Exit Function
        Case 401
            vla_step_text = "Put ""m4"" into cell A87. [line 736]" ' vla:4091
            Exit Function
        Case 402
            vla_step_text = "Put ""m5"" into cell A88. [line 737]" ' vla:4092
            Exit Function
        Case 403
            vla_step_text = "Delete A84:A86 and shift cells up. [line 738]" ' vla:4093
            Exit Function
        Case 404
            vla_step_text = "Put ""x1"" into cell B90. [line 740]" ' vla:4094
            Exit Function
        Case 405
            vla_step_text = "Put ""x2"" into cell C90. [line 741]" ' vla:4095
            Exit Function
        Case 406
            vla_step_text = "Put ""x3"" into cell D90. [line 742]" ' vla:4096
            Exit Function
        Case 407
            vla_step_text = "Put ""rightdata"" into cell E90. [line 743]" ' vla:4097
            Exit Function
        Case 408
            vla_step_text = "Delete B90:D90 and shift cells left. [line 744]" ' vla:4098
            Exit Function
        Case 409
            vla_step_text = "Put 1 into cell A94. [line 746]" ' vla:4099
            Exit Function
        Case 410
            vla_step_text = "Put 2 into cell B94. [line 747]" ' vla:4100
            Exit Function
        Case 411
            vla_step_text = "Put 3 into cell A95. [line 748]" ' vla:4101
            Exit Function
        Case 412
            vla_step_text = "Put 5 into cell A96. [line 749]" ' vla:4102
            Exit Function
        Case 413
            vla_step_text = "Put 6 into cell B96. [line 750]" ' vla:4103
            Exit Function
        Case 414
            vla_step_text = "Delete blank rows in A94:B96. [line 751]" ' vla:4104
            Exit Function
        Case 415
            vla_step_text = "Put 1 into cell A150. [line 754]" ' vla:4105
            Exit Function
        Case 416
            vla_step_text = "Put 2 into cell A151. [line 755]" ' vla:4106
            Exit Function
        Case 417
            vla_step_text = "Put 3 into cell A152. [line 756]" ' vla:4107
            Exit Function
        Case 418
            vla_step_text = "Put 4 into cell A153. [line 757]" ' vla:4108
            Exit Function
        Case 419
            vla_step_text = "Band every other row of A150:D153 ""#D9D9D9"". [line 758]" ' vla:4109
            Exit Function
        Case 420
            vla_step_text = "Make row 165 a header row. [line 759]" ' vla:4110
            Exit Function
        Case 421
            vla_step_text = "Fill A170:A179 with a series starting at 1. [line 762]" ' vla:4111
            Exit Function
        Case 422
            vla_step_text = "Fill A180:A184 with a series starting at 1 with step 2. [line 763]" ' vla:4112
            Exit Function
        Case 423
            vla_step_text = "Fill A190:A194 with a growth series starting at 2. [line 764]" ' vla:4113
            Exit Function
        Case 424
            vla_step_text = "Fill A200:A203 with a growth series starting at 2 with step 3. [line 765]" ' vla:4114
            Exit Function
        Case 425
            vla_step_text = "Make cell Z500 bold. [line 777]" ' vla:4115
            Exit Function
        Case 426
            vla_step_text = "Clear everything from Z500. [line 778]" ' vla:4116
            Exit Function
        Case 427
            vla_step_text = "Remove trailing empty rows and columns. [line 779]" ' vla:4117
            Exit Function
        Case 428
            vla_step_text = "Try: [line 792]" ' vla:4118
            Exit Function
        Case 429
            vla_step_text = "Delete sheet GFormat. [line 793]" ' vla:4119
            Exit Function
        Case 430
            vla_step_text = "Work on sheet GFormat. [line 795]" ' vla:4120
            Exit Function
        Case 431
            vla_step_text = "Make range A1:C1 bold. [line 799]" ' vla:4121
            Exit Function
        Case 432
            vla_step_text = "Make range A2:C2 italic. [line 800]" ' vla:4122
            Exit Function
        Case 433
            vla_step_text = "Make range A3:C3 blue. [line 801]" ' vla:4123
            Exit Function
        Case 434
            vla_step_text = "Set font-color of range A4:C4 to ""#FF0000"". [line 802]" ' vla:4124
            Exit Function
        Case 435
            vla_step_text = "Set fill-color of range A5:C5 to ""#00FF00"". [line 803]" ' vla:4125
            Exit Function
        Case 436
            vla_step_text = "Make range A6:C6 yellow. [line 804]" ' vla:4126
            Exit Function
        Case 437
            vla_step_text = "Clear fill-color of range B6:C6. [line 805]" ' vla:4127
            Exit Function
        Case 438
            vla_step_text = "Set font size of range A7:C7 to 16. [line 806]" ' vla:4128
            Exit Function
        Case 439
            vla_step_text = "Make range A8:C8 bold. [line 809]" ' vla:4129
            Exit Function
        Case 440
            vla_step_text = "Make range B8:C8 not bold. [line 810]" ' vla:4130
            Exit Function
        Case 441
            vla_step_text = "Make range A9:C9 italic. [line 811]" ' vla:4131
            Exit Function
        Case 442
            vla_step_text = "Make cell C9 not italic. [line 812]" ' vla:4132
            Exit Function
        Case 443
            vla_step_text = "Underline range A10:C10. [line 813]" ' vla:4133
            Exit Function
        Case 444
            vla_step_text = "Strike through range A11:C11. [line 814]" ' vla:4134
            Exit Function
        Case 445
            vla_step_text = "Underline range A12:C12. [line 815]" ' vla:4135
            Exit Function
        Case 446
            vla_step_text = "Remove underline from range B12:C12. [line 816]" ' vla:4136
            Exit Function
        Case 447
            vla_step_text = "Strike through range A13:C13. [line 817]" ' vla:4137
            Exit Function
        Case 448
            vla_step_text = "Remove the strikethrough from cell C13. [line 818]" ' vla:4138
            Exit Function
        Case 449
            vla_step_text = "Set font of range A14:C14 to ""Courier New"". [line 819]" ' vla:4139
            Exit Function
        Case 450
            vla_step_text = "Align range A15:C15 to the top. [line 823]" ' vla:4140
            Exit Function
        Case 451
            vla_step_text = "Align range A16:C16 to the middle. [line 824]" ' vla:4141
            Exit Function
        Case 452
            vla_step_text = "Align range A17:C17 to the top. [line 825]" ' vla:4142
            Exit Function
        Case 453
            vla_step_text = "Align range B17:C17 to the bottom. [line 826]" ' vla:4143
            Exit Function
        Case 454
            vla_step_text = "Indent range A18:C18 by 2. [line 827]" ' vla:4144
            Exit Function
        Case 455
            vla_step_text = "Rotate text in range A19:C19 by 45 degrees. [line 828]" ' vla:4145
            Exit Function
        Case 456
            vla_step_text = "Add a border around range B21:D23. [line 834]" ' vla:4146
            Exit Function
        Case 457
            vla_step_text = "Add a bottom border to range B25:D25. [line 835]" ' vla:4147
            Exit Function
        Case 458
            vla_step_text = "Add borders to every cell in range B27:D29. [line 836]" ' vla:4148
            Exit Function
        Case 459
            vla_step_text = "Add borders to every cell in range B31:D33. [line 837]" ' vla:4149
            Exit Function
        Case 460
            vla_step_text = "Remove borders from range C31:D33. [line 838]" ' vla:4150
            Exit Function
        Case 461
            vla_step_text = "Set width of column F to 40. [line 846]" ' vla:4151
            Exit Function
        Case 462
            vla_step_text = "Put 1234.56 into cell F40. [line 847]" ' vla:4152
            Exit Function
        Case 463
            vla_step_text = "Format cell F40 as a number. [line 848]" ' vla:4153
            Exit Function
        Case 464
            vla_step_text = "Put 1234.56 into cell F41. [line 849]" ' vla:4154
            Exit Function
        Case 465
            vla_step_text = "Format cell F41 as a number with 3 decimals. [line 850]" ' vla:4155
            Exit Function
        Case 466
            vla_step_text = "Put 1234.56 into cell F42. [line 851]" ' vla:4156
            Exit Function
        Case 467
            vla_step_text = "Format cell F42 as a number with thousands separators. [line 852]" ' vla:4157
            Exit Function
        Case 468
            vla_step_text = "Put 1234.56 into cell F43. [line 853]" ' vla:4158
            Exit Function
        Case 469
            vla_step_text = "Format cell F43 as a number with 0 decimals and thousands separators. [line 854]" ' vla:4159
            Exit Function
        Case 470
            vla_step_text = "Put 1234.56 into cell F44. [line 855]" ' vla:4160
            Exit Function
        Case 471
            vla_step_text = "Format cell F44 as dollars. [line 856]" ' vla:4161
            Exit Function
        Case 472
            vla_step_text = "Put 1234.56 into cell F45. [line 857]" ' vla:4162
            Exit Function
        Case 473
            vla_step_text = "Format cell F45 as euros with 0 decimals. [line 858]" ' vla:4163
            Exit Function
        Case 474
            vla_step_text = "Put 1234.56 into cell F46. [line 859]" ' vla:4164
            Exit Function
        Case 475
            vla_step_text = "Format cell F46 as pounds. [line 860]" ' vla:4165
            Exit Function
        Case 476
            vla_step_text = "Put 1234.56 into cell F47. [line 861]" ' vla:4166
            Exit Function
        Case 477
            vla_step_text = "Format cell F47 as accounting in dollars. [line 862]" ' vla:4167
            Exit Function
        Case 478
            vla_step_text = "Put -1234.56 into cell F48. [line 863]" ' vla:4168
            Exit Function
        Case 479
            vla_step_text = "Format cell F48 as accounting in euros with 0 decimals. [line 864]" ' vla:4169
            Exit Function
        Case 480
            vla_step_text = "Put 0.125 into cell F49. [line 865]" ' vla:4170
            Exit Function
        Case 481
            vla_step_text = "Format range F49 as percent. [line 866]" ' vla:4171
            Exit Function
        Case 482
            vla_step_text = "Put 0.125 into cell F50. [line 867]" ' vla:4172
            Exit Function
        Case 483
            vla_step_text = "Format cell F50 as percent with 2 decimals. [line 868]" ' vla:4173
            Exit Function
        Case 484
            vla_step_text = "Put 46000 into cell F51. [line 869]" ' vla:4174
            Exit Function
        Case 485
            vla_step_text = "Format cell F51 as a short date. [line 870]" ' vla:4175
            Exit Function
        Case 486
            vla_step_text = "Put 46000 into cell F52. [line 871]" ' vla:4176
            Exit Function
        Case 487
            vla_step_text = "Format cell F52 as a long date. [line 872]" ' vla:4177
            Exit Function
        Case 488
            vla_step_text = "Put 46000 into cell F53. [line 873]" ' vla:4178
            Exit Function
        Case 489
            vla_step_text = "Format cell F53 as an ISO date. [line 874]" ' vla:4179
            Exit Function
        Case 490
            vla_step_text = "Put 0.5625 into cell F54. [line 875]" ' vla:4180
            Exit Function
        Case 491
            vla_step_text = "Format cell F54 as a time. [line 876]" ' vla:4181
            Exit Function
        Case 492
            vla_step_text = "Put 42 into cell F55. [line 877]" ' vla:4182
            Exit Function
        Case 493
            vla_step_text = "Format cell F55 as text for new entries. [line 878]" ' vla:4183
            Exit Function
        Case 494
            vla_step_text = "Put 1234.56 into cell F56. [line 879]" ' vla:4184
            Exit Function
        Case 495
            vla_step_text = "Format cell F56 as dollars. [line 880]" ' vla:4185
            Exit Function
        Case 496
            vla_step_text = "Format cell F56 as general. [line 881]" ' vla:4186
            Exit Function
        Case 497
            vla_step_text = "Put 42 into cell F57. [line 882]" ' vla:4187
            Exit Function
        Case 498
            vla_step_text = "Format cell F57 using ""00000"". [line 883]" ' vla:4188
            Exit Function
        Case 499
            vla_step_text = "Add a border colored ""#FF0000"" around range H40:J42. [line 889]" ' vla:4189
            Exit Function
        Case 500
            vla_step_text = "Add a bottom border colored green to range H44:J44. [line 890]" ' vla:4190
            Exit Function
        Case 501
            vla_step_text = "Add borders colored blue to every cell in range H46:J48. [line 891]" ' vla:4191
            Exit Function
        Case 502
            vla_step_text = "Try: [line 899]" ' vla:4192
            Exit Function
        Case 503
            vla_step_text = "Delete sheet GSortFilter. [line 900]" ' vla:4193
            Exit Function
        Case 504
            vla_step_text = "Work on sheet GSortFilter. [line 902]" ' vla:4194
            Exit Function
        Case 505
            vla_step_text = "Put ""Item"" into cell A1. [line 906]" ' vla:4195
            Exit Function
        Case 506
            vla_step_text = "Put ""Qty"" into cell B1. [line 907]" ' vla:4196
            Exit Function
        Case 507
            vla_step_text = "Put ""a"" into cell A2. [line 908]" ' vla:4197
            Exit Function
        Case 508
            vla_step_text = "Put 2 into cell B2. [line 909]" ' vla:4198
            Exit Function
        Case 509
            vla_step_text = "Put ""b"" into cell A3. [line 910]" ' vla:4199
            Exit Function
        Case 510
            vla_step_text = "Put 9 into cell B3. [line 911]" ' vla:4200
            Exit Function
        Case 511
            vla_step_text = "Put ""c"" into cell A4. [line 912]" ' vla:4201
            Exit Function
        Case 512
            vla_step_text = "Put 5 into cell B4. [line 913]" ' vla:4202
            Exit Function
        Case 513
            vla_step_text = "Sort this sheet by column B descending with a header row. [line 914]" ' vla:4203
            Exit Function
        Case 514
            vla_step_text = "Put ""Name"" into cell D1. [line 918]" ' vla:4204
            Exit Function
        Case 515
            vla_step_text = "Put ""Score"" into cell E1. [line 919]" ' vla:4205
            Exit Function
        Case 516
            vla_step_text = "Put ""p"" into cell D2. [line 920]" ' vla:4206
            Exit Function
        Case 517
            vla_step_text = "Put 3 into cell E2. [line 921]" ' vla:4207
            Exit Function
        Case 518
            vla_step_text = "Put ""q"" into cell D3. [line 922]" ' vla:4208
            Exit Function
        Case 519
            vla_step_text = "Put 1 into cell E3. [line 923]" ' vla:4209
            Exit Function
        Case 520
            vla_step_text = "Put ""r"" into cell D4. [line 924]" ' vla:4210
            Exit Function
        Case 521
            vla_step_text = "Put 2 into cell E4. [line 925]" ' vla:4211
            Exit Function
        Case 522
            vla_step_text = "Sort range D1:E4 by column E with a header row. [line 926]" ' vla:4212
            Exit Function
        Case 523
            vla_step_text = "Put ""x"" into cell G2. [line 929]" ' vla:4213
            Exit Function
        Case 524
            vla_step_text = "Put 5 into cell H2. [line 930]" ' vla:4214
            Exit Function
        Case 525
            vla_step_text = "Put ""y"" into cell G3. [line 931]" ' vla:4215
            Exit Function
        Case 526
            vla_step_text = "Put 9 into cell H3. [line 932]" ' vla:4216
            Exit Function
        Case 527
            vla_step_text = "Put ""z"" into cell G4. [line 933]" ' vla:4217
            Exit Function
        Case 528
            vla_step_text = "Put 7 into cell H4. [line 934]" ' vla:4218
            Exit Function
        Case 529
            vla_step_text = "Sort range G2:H4 by column H descending without a header row. [line 935]" ' vla:4219
            Exit Function
        Case 530
            vla_step_text = "Put ""Region"" into cell J1. [line 939]" ' vla:4220
            Exit Function
        Case 531
            vla_step_text = "Put ""Amount"" into cell K1. [line 940]" ' vla:4221
            Exit Function
        Case 532
            vla_step_text = "Put ""Tag"" into cell L1. [line 941]" ' vla:4222
            Exit Function
        Case 533
            vla_step_text = "Put ""West"" into cell J2. [line 942]" ' vla:4223
            Exit Function
        Case 534
            vla_step_text = "Put 100 into cell K2. [line 943]" ' vla:4224
            Exit Function
        Case 535
            vla_step_text = "Put ""a"" into cell L2. [line 944]" ' vla:4225
            Exit Function
        Case 536
            vla_step_text = "Put ""East"" into cell J3. [line 945]" ' vla:4226
            Exit Function
        Case 537
            vla_step_text = "Put 50 into cell K3. [line 946]" ' vla:4227
            Exit Function
        Case 538
            vla_step_text = "Put ""b"" into cell L3. [line 947]" ' vla:4228
            Exit Function
        Case 539
            vla_step_text = "Put ""West"" into cell J4. [line 948]" ' vla:4229
            Exit Function
        Case 540
            vla_step_text = "Put 300 into cell K4. [line 949]" ' vla:4230
            Exit Function
        Case 541
            vla_step_text = "Put ""c"" into cell L4. [line 950]" ' vla:4231
            Exit Function
        Case 542
            vla_step_text = "Put ""East"" into cell J5. [line 951]" ' vla:4232
            Exit Function
        Case 543
            vla_step_text = "Put 250 into cell K5. [line 952]" ' vla:4233
            Exit Function
        Case 544
            vla_step_text = "Put ""d"" into cell L5. [line 953]" ' vla:4234
            Exit Function
        Case 545
            vla_step_text = "Sort range J1:L5 by column J ascending then by column K descending with a header row. [line 954]" ' vla:4235
            Exit Function
        Case 546
            vla_step_text = "Put 2 into cell N2. [line 957]" ' vla:4236
            Exit Function
        Case 547
            vla_step_text = "Put 9 into cell O2. [line 958]" ' vla:4237
            Exit Function
        Case 548
            vla_step_text = "Put 1 into cell N3. [line 959]" ' vla:4238
            Exit Function
        Case 549
            vla_step_text = "Put 8 into cell O3. [line 960]" ' vla:4239
            Exit Function
        Case 550
            vla_step_text = "Put 2 into cell N4. [line 961]" ' vla:4240
            Exit Function
        Case 551
            vla_step_text = "Put 7 into cell O4. [line 962]" ' vla:4241
            Exit Function
        Case 552
            vla_step_text = "Sort range N2:O4 by column N then by column O without a header row. [line 963]" ' vla:4242
            Exit Function
        Case 553
            vla_step_text = "Put ""Region"" into cell A40. [line 972]" ' vla:4243
            Exit Function
        Case 554
            vla_step_text = "Put ""Amount"" into cell B40. [line 973]" ' vla:4244
            Exit Function
        Case 555
            vla_step_text = "Put ""West"" into cell A41. [line 974]" ' vla:4245
            Exit Function
        Case 556
            vla_step_text = "Put 1 into cell B41. [line 975]" ' vla:4246
            Exit Function
        Case 557
            vla_step_text = "Put ""East"" into cell A42. [line 976]" ' vla:4247
            Exit Function
        Case 558
            vla_step_text = "Put 2 into cell B42. [line 977]" ' vla:4248
            Exit Function
        Case 559
            vla_step_text = "Put ""West"" into cell A43. [line 978]" ' vla:4249
            Exit Function
        Case 560
            vla_step_text = "Put 3 into cell B43. [line 979]" ' vla:4250
            Exit Function
        Case 561
            vla_step_text = "Filter range A40:B43 to show rows where column A is ""West"". [line 980]" ' vla:4251
            Exit Function
        Case 562
            vla_step_text = "Copy only the visible cells of range A40:B43 to D46. [line 981]" ' vla:4252
            Exit Function
        Case 563
            vla_step_text = "Remove the filters. [line 982]" ' vla:4253
            Exit Function
        Case 564
            vla_step_text = "Try: [line 988]" ' vla:4254
            Exit Function
        Case 565
            vla_step_text = "Filter range A40:B43 to show rows where column B is greater than ""abc"". [line 989]" ' vla:4255
            Exit Function
        Case 566
            vla_step_text = "Put ""Region"" into cell A20. [line 994]" ' vla:4256
            Exit Function
        Case 567
            vla_step_text = "Put ""Amount"" into cell B20. [line 995]" ' vla:4257
            Exit Function
        Case 568
            vla_step_text = "Put ""Code"" into cell C20. [line 996]" ' vla:4258
            Exit Function
        Case 569
            vla_step_text = "Put ""West"" into cell A21. [line 997]" ' vla:4259
            Exit Function
        Case 570
            vla_step_text = "Put 100 into cell B21. [line 998]" ' vla:4260
            Exit Function
        Case 571
            vla_step_text = "Put ""5*3"" into cell C21. [line 999]" ' vla:4261
            Exit Function
        Case 572
            vla_step_text = "Put ""East"" into cell A22. [line 1000]" ' vla:4262
            Exit Function
        Case 573
            vla_step_text = "Put 250 into cell B22. [line 1001]" ' vla:4263
            Exit Function
        Case 574
            vla_step_text = "Put ""53"" into cell C22. [line 1002]" ' vla:4264
            Exit Function
        Case 575
            vla_step_text = "Put ""Western"" into cell A23. [line 1003]" ' vla:4265
            Exit Function
        Case 576
            vla_step_text = "Put 100 into cell B23. [line 1004]" ' vla:4266
            Exit Function
        Case 577
            vla_step_text = "Put ""x"" into cell C23. [line 1005]" ' vla:4267
            Exit Function
        Case 578
            vla_step_text = "Put ""West"" into cell A24. [line 1006]" ' vla:4268
            Exit Function
        Case 579
            vla_step_text = "Put 50 into cell B24. [line 1007]" ' vla:4269
            Exit Function
        Case 580
            vla_step_text = "Put ""y"" into cell C24. [line 1008]" ' vla:4270
            Exit Function
        Case 581
            vla_step_text = "Put ""east"" into cell A25. [line 1009]" ' vla:4271
            Exit Function
        Case 582
            vla_step_text = "Put 300 into cell B25. [line 1010]" ' vla:4272
            Exit Function
        Case 583
            vla_step_text = "Put ""z"" into cell C25. [line 1011]" ' vla:4273
            Exit Function
        Case 584
            vla_step_text = "Put ""North"" into cell A26. [line 1012]" ' vla:4274
            Exit Function
        Case 585
            vla_step_text = "Put 175 into cell B26. [line 1013]" ' vla:4275
            Exit Function
        Case 586
            vla_step_text = "Put ""w"" into cell C26. [line 1014]" ' vla:4276
            Exit Function
        Case 587
            vla_step_text = "Put ""East"" into cell A27. [line 1015]" ' vla:4277
            Exit Function
        Case 588
            vla_step_text = "Put 75 into cell B27. [line 1016]" ' vla:4278
            Exit Function
        Case 589
            vla_step_text = "Put ""v"" into cell C27. [line 1017]" ' vla:4279
            Exit Function
        Case 590
            vla_step_text = "Format range B21:B27 as dollars. [line 1018]" ' vla:4280
            Exit Function
        Case 591
            vla_step_text = "Filter range A20:C27 to show rows where column A is ""West"". [line 1020]" ' vla:4281
            Exit Function
        Case 592
            vla_step_text = "Copy only the visible cells of range A20:C27 to E30. [line 1021]" ' vla:4282
            Exit Function
        Case 593
            vla_step_text = "Clear the filter conditions. [line 1022]" ' vla:4283
            Exit Function
        Case 594
            vla_step_text = "Filter range A20:C27 to show rows where column B is 100. [line 1024]" ' vla:4284
            Exit Function
        Case 595
            vla_step_text = "Copy only the visible cells of range A20:C27 to I30. [line 1025]" ' vla:4285
            Exit Function
        Case 596
            vla_step_text = "Clear the filter conditions. [line 1026]" ' vla:4286
            Exit Function
        Case 597
            vla_step_text = "Filter range A20:C27 to show rows where column C contains ""*"". [line 1028]" ' vla:4287
            Exit Function
        Case 598
            vla_step_text = "Copy only the visible cells of range A20:C27 to M30. [line 1029]" ' vla:4288
            Exit Function
        Case 599
            vla_step_text = "Clear the filter conditions. [line 1030]" ' vla:4289
            Exit Function
        Case 600
            vla_step_text = "Filter range A20:C27 to show rows where column B is greater than 150. [line 1033]" ' vla:4290
            Exit Function
        Case 601
            vla_step_text = "Filter range A20:C27 to show rows where column A is ""East"". [line 1034]" ' vla:4291
            Exit Function
        Case 602
            vla_step_text = "Copy only the visible cells of range A20:C27 to Q30. [line 1035]" ' vla:4292
            Exit Function
        Case 603
            vla_step_text = "Clear the filter conditions. [line 1036]" ' vla:4293
            Exit Function
        Case 604
            vla_step_text = "Filter range A20:C27 to show rows where column B is less than 99.5. [line 1041]" ' vla:4294
            Exit Function
        Case 605
            vla_step_text = "Copy only the visible cells of range A20:C27 to U30. [line 1042]" ' vla:4295
            Exit Function
        Case 606
            vla_step_text = "Add filters to range A20:C27. [line 1049]" ' vla:4296
            Exit Function
        Case 607
            vla_step_text = "Clear the filter conditions. [line 1050]" ' vla:4297
            Exit Function
        Case 608
            vla_step_text = "Try: [line 1051]" ' vla:4298
            Exit Function
        Case 609
            vla_step_text = "Add filters to range A40:B43. [line 1052]" ' vla:4299
            Exit Function
        Case 610
            vla_step_text = "Try: [line 1060]" ' vla:4300
            Exit Function
        Case 611
            vla_step_text = "Delete sheet GText. [line 1061]" ' vla:4301
            Exit Function
        Case 612
            vla_step_text = "Work on sheet GText. [line 1063]" ' vla:4302
            Exit Function
        Case 613
            vla_step_text = "Put ""widget"" into cell A1. [line 1067]" ' vla:4303
            Exit Function
        Case 614
            vla_step_text = "Put 42 into cell A2. [line 1068]" ' vla:4304
            Exit Function
        Case 615
            vla_step_text = "Put formula ""=CHAR(97)&CHAR(98)"" into cell A3. [line 1069]" ' vla:4305
            Exit Function
        Case 616
            vla_step_text = "Put ""'true"" into cell A4. [line 1070]" ' vla:4306
            Exit Function
        Case 617
            vla_step_text = "Put ""'=abc"" into cell A5. [line 1071]" ' vla:4307
            Exit Function
        Case 618
            vla_step_text = "Make range A1:A5 upper case. [line 1072]" ' vla:4308
            Exit Function
        Case 619
            vla_step_text = "Put ""HELLO World"" into cell B1. [line 1075]" ' vla:4309
            Exit Function
        Case 620
            vla_step_text = "Put ""MiXeD"" into cell B2. [line 1076]" ' vla:4310
            Exit Function
        Case 621
            vla_step_text = "Make column B lowercase. [line 1077]" ' vla:4311
            Exit Function
        Case 622
            vla_step_text = "Put ""don't stop"" into cell C1. [line 1080]" ' vla:4312
            Exit Function
        Case 623
            vla_step_text = "Put ""3rd quarter"" into cell C2. [line 1081]" ' vla:4313
            Exit Function
        Case 624
            vla_step_text = "Put ""o'neil"" into cell C3. [line 1082]" ' vla:4314
            Exit Function
        Case 625
            vla_step_text = "Capitalize each word in range C1:C3 after any space. [line 1083]" ' vla:4315
            Exit Function
        Case 626
            vla_step_text = "Put ""don't stop"" into cell D1. [line 1084]" ' vla:4316
            Exit Function
        Case 627
            vla_step_text = "Put ""3rd quarter"" into cell D2. [line 1085]" ' vla:4317
            Exit Function
        Case 628
            vla_step_text = "Put ""o'neil"" into cell D3. [line 1086]" ' vla:4318
            Exit Function
        Case 629
            vla_step_text = "Capitalize each word in column D after any non-letter. [line 1087]" ' vla:4319
            Exit Function
        Case 630
            vla_step_text = "Put ""  a   b  "" into cell E1. [line 1092]" ' vla:4320
            Exit Function
        Case 631
            vla_step_text = "Put ""'  00123 "" into cell E2. [line 1093]" ' vla:4321
            Exit Function
        Case 632
            vla_step_text = "Put ""   "" into cell E3. [line 1094]" ' vla:4322
            Exit Function
        Case 633
            vla_step_text = "Put formula ""=UNICHAR(160)&CHAR(120)&UNICHAR(160)&UNICHAR(160)&CHAR(121)"" into cell E4. [line 1095]" ' vla:4323
            Exit Function
        Case 634
            vla_step_text = "Convert range E4:E4 to values. [line 1096]" ' vla:4324
            Exit Function
        Case 635
            vla_step_text = "Remove extra spaces from range E1:E4. [line 1097]" ' vla:4325
            Exit Function
        Case 636
            vla_step_text = "Put formula ""=CHAR(97)&CHAR(10)&CHAR(98)"" into cell F1. [line 1100]" ' vla:4326
            Exit Function
        Case 637
            vla_step_text = "Put formula ""=CHAR(99)&CHAR(9)&CHAR(100)"" into cell F2. [line 1101]" ' vla:4327
            Exit Function
        Case 638
            vla_step_text = "Convert range F1:F2 to values. [line 1102]" ' vla:4328
            Exit Function
        Case 639
            vla_step_text = "Remove non-printing characters from range F1:F2. [line 1103]" ' vla:4329
            Exit Function
        Case 640
            vla_step_text = "Put ""Widget"" into cell G1. [line 1108]" ' vla:4330
            Exit Function
        Case 641
            vla_step_text = "Put ""banana"" into cell G2. [line 1109]" ' vla:4331
            Exit Function
        Case 642
            vla_step_text = "Set gtext-row to row of ""Widget"" in column G. [line 1110]" ' vla:4332
            Exit Function
        Case 643
            vla_step_text = "Replace ""AN"" with ""xy"" in range G2:G2. [line 1111]" ' vla:4333
            Exit Function
        Case 644
            vla_step_text = "Format range I1:I20 as text for new entries. [line 1115]" ' vla:4334
            Exit Function
        Case 645
            vla_step_text = "Set gtext-code to ""INV-ab-Cd"". [line 1116]" ' vla:4335
            Exit Function
        Case 646
            vla_step_text = "Set gtext-part to the text before ""-"" in gtext-code. [line 1117]" ' vla:4336
            Exit Function
        Case 647
            vla_step_text = "Put gtext-part into cell I1. [line 1118]" ' vla:4337
            Exit Function
        Case 648
            vla_step_text = "Set gtext-part to the text after ""-"" in gtext-code. [line 1119]" ' vla:4338
            Exit Function
        Case 649
            vla_step_text = "Put gtext-part into cell I2. [line 1120]" ' vla:4339
            Exit Function
        Case 650
            vla_step_text = "Set gtext-part to the text after the last ""-"" in gtext-code. [line 1121]" ' vla:4340
            Exit Function
        Case 651
            vla_step_text = "Put gtext-part into cell I3. [line 1122]" ' vla:4341
            Exit Function
        Case 652
            vla_step_text = "Set gtext-part to the text before ""c"" in gtext-code. [line 1125]" ' vla:4342
            Exit Function
        Case 653
            vla_step_text = "Put gtext-part into cell I4. [line 1126]" ' vla:4343
            Exit Function
        Case 654
            vla_step_text = "Set gtext-part to the first 3 characters of gtext-code. [line 1127]" ' vla:4344
            Exit Function
        Case 655
            vla_step_text = "Put gtext-part into cell I5. [line 1128]" ' vla:4345
            Exit Function
        Case 656
            vla_step_text = "Set gtext-part to the last 2 characters of gtext-code. [line 1129]" ' vla:4346
            Exit Function
        Case 657
            vla_step_text = "Put gtext-part into cell I6. [line 1130]" ' vla:4347
            Exit Function
        Case 658
            vla_step_text = "Set gtext-part to 42 padded on the left with ""0"" to 5 characters. [line 1132]" ' vla:4348
            Exit Function
        Case 659
            vla_step_text = "Put gtext-part into cell I7. [line 1133]" ' vla:4349
            Exit Function
        Case 660
            vla_step_text = "Set gtext-part to ""ab"" padded on the right with ""."" to 4 characters. [line 1134]" ' vla:4350
            Exit Function
        Case 661
            vla_step_text = "Put gtext-part into cell I8. [line 1135]" ' vla:4351
            Exit Function
        Case 662
            vla_step_text = "Put ""x"" into cell J1. [line 1138]" ' vla:4352
            Exit Function
        Case 663
            vla_step_text = "Put 7 into cell J3. [line 1139]" ' vla:4353
            Exit Function
        Case 664
            vla_step_text = "Put ""y"" into cell J4. [line 1140]" ' vla:4354
            Exit Function
        Case 665
            vla_step_text = "Set gtext-list to range J1:J4 as one list. [line 1141]" ' vla:4355
            Exit Function
        Case 666
            vla_step_text = "Put gtext-list into cell I9. [line 1142]" ' vla:4356
            Exit Function
        Case 667
            vla_step_text = "Set gtext-list to range J1:J4 as one list separated by ""; "". [line 1143]" ' vla:4357
            Exit Function
        Case 668
            vla_step_text = "Put gtext-list into cell I10. [line 1144]" ' vla:4358
            Exit Function
        Case 669
            vla_step_text = "Set gtext-list to column J as one list. [line 1145]" ' vla:4359
            Exit Function
        Case 670
            vla_step_text = "Put gtext-list into cell I11. [line 1146]" ' vla:4360
            Exit Function
        Case 671
            vla_step_text = "Set gtext-part to ""  a   b "" with extra spaces removed. [line 1150]" ' vla:4361
            Exit Function
        Case 672
            vla_step_text = "Put gtext-part into cell I12. [line 1151]" ' vla:4362
            Exit Function
        Case 673
            vla_step_text = "Set gtext-part to ""o'neil"" with each word capitalized after any non-letter. [line 1152]" ' vla:4363
            Exit Function
        Case 674
            vla_step_text = "Put gtext-part into cell I13. [line 1153]" ' vla:4364
            Exit Function
        Case 675
            vla_step_text = "Set gtext-part to ""don't stop"" with each word capitalized after any space. [line 1154]" ' vla:4365
            Exit Function
        Case 676
            vla_step_text = "Put gtext-part into cell I14. [line 1155]" ' vla:4366
            Exit Function
        Case 677
            vla_step_text = "Put formula ""=CHAR(112)&CHAR(10)&CHAR(113)"" into cell K1. [line 1156]" ' vla:4367
            Exit Function
        Case 678
            vla_step_text = "Convert range K1:K1 to values. [line 1157]" ' vla:4368
            Exit Function
        Case 679
            vla_step_text = "Set gtext-part to cell K1 with non-printing characters removed. [line 1158]" ' vla:4369
            Exit Function
        Case 680
            vla_step_text = "Put gtext-part into cell I15. [line 1159]" ' vla:4370
            Exit Function
        Case 681
            vla_step_text = "Try: [line 1164]" ' vla:4371
            Exit Function
        Case 682
            vla_step_text = "Set gtext-part to the text before ""#"" in gtext-code. [line 1165]" ' vla:4372
            Exit Function
        Case 683
            vla_step_text = "Put gtext-part into cell I16. [line 1167]" ' vla:4373
            Exit Function
        Case 684
            vla_step_text = "Put ""Nan"" into cell L1. [line 1173]" ' vla:4374
            Exit Function
        Case 685
            vla_step_text = "Put formula ""=TAN(0)"" into cell L2. [line 1174]" ' vla:4375
            Exit Function
        Case 686
            vla_step_text = "Replace ""an"" with ""xy"" in range L1:L2. [line 1175]" ' vla:4376
            Exit Function
        Case 687
            vla_step_text = "Put ""Q"" into cell L3. [line 1176]" ' vla:4377
            Exit Function
        Case 688
            vla_step_text = "Replace ""Q"" with ""=1+1"" in range L3:L3. [line 1177]" ' vla:4378
            Exit Function
        Case 689
            vla_step_text = "Put ""N/A"" into cell L4. [line 1178]" ' vla:4379
            Exit Function
        Case 690
            vla_step_text = "Replace ""N/A"" with 0 in range L4:L4. [line 1179]" ' vla:4380
            Exit Function
        Case 691
            vla_step_text = "Put formula ""=ABS(-2)"" into cell L5. [line 1182]" ' vla:4381
            Exit Function
        Case 692
            vla_step_text = "Put ""ABS"" into cell L6. [line 1183]" ' vla:4382
            Exit Function
        Case 693
            vla_step_text = "Replace ""ABS"" with ""SIGN"" in the formulas of range L5:L6. [line 1184]" ' vla:4383
            Exit Function
        Case 694
            vla_step_text = "Put ""zqz"" into cell L7. [line 1187]" ' vla:4384
            Exit Function
        Case 695
            vla_step_text = "Replace ""zqz"" with ""done"" on this sheet. [line 1188]" ' vla:4385
            Exit Function
        Case 696
            vla_step_text = "Put formula ""=ABS(-3)"" into cell L8. [line 1189]" ' vla:4386
            Exit Function
        Case 697
            vla_step_text = "Replace ""-3"" with ""-4"" in the formulas on this sheet. [line 1190]" ' vla:4387
            Exit Function
        Case 698
            vla_step_text = "Put ""x,0042,7"" into cell N1. [line 1195]" ' vla:4388
            Exit Function
        Case 699
            vla_step_text = "Split column N by "","" as text. [line 1196]" ' vla:4389
            Exit Function
        Case 700
            vla_step_text = "Put ""y,0042,7,-3.5"" into cell R1. [line 1197]" ' vla:4390
            Exit Function
        Case 701
            vla_step_text = "Split column R by "","" reading numbers. [line 1198]" ' vla:4391
            Exit Function
        Case 702
            vla_step_text = "Put ""p,q"" into cell W1. [line 1199]" ' vla:4392
            Exit Function
        Case 703
            vla_step_text = "Put ""keep"" into cell X1. [line 1200]" ' vla:4393
            Exit Function
        Case 704
            vla_step_text = "Try: [line 1201]" ' vla:4394
            Exit Function
        Case 705
            vla_step_text = "Split column W by "","" as text. [line 1202]" ' vla:4395
            Exit Function
        Case 706
            vla_step_text = "Put ""apple"" into cell Z1. [line 1205]" ' vla:4396
            Exit Function
        Case 707
            vla_step_text = "Put ""Banana"" into cell AA1. [line 1206]" ' vla:4397
            Exit Function
        Case 708
            vla_step_text = "Put ""cherry"" into cell Z2. [line 1207]" ' vla:4398
            Exit Function
        Case 709
            vla_step_text = "Put ""banana split"" into cell AA2. [line 1208]" ' vla:4399
            Exit Function
        Case 710
            vla_step_text = "Put 7 into cell Z3. [line 1209]" ' vla:4400
            Exit Function
        Case 711
            vla_step_text = "Put ""BANANA"" into cell AA3. [line 1210]" ' vla:4401
            Exit Function
        Case 712
            vla_step_text = "Set gtext-found to the row of the first cell in range Z1:AA3 containing ""nan"". [line 1211]" ' vla:4402
            Exit Function
        Case 713
            vla_step_text = "Put gtext-found into cell L9. [line 1212]" ' vla:4403
            Exit Function
        Case 714
            vla_step_text = "Set gtext-found to the column of the first cell in range Z1:AA3 containing ""nan"". [line 1213]" ' vla:4404
            Exit Function
        Case 715
            vla_step_text = "Put gtext-found into cell L10. [line 1214]" ' vla:4405
            Exit Function
        Case 716
            vla_step_text = "Set gtext-found to the row of the first cell in column Z containing ""err"". [line 1215]" ' vla:4406
            Exit Function
        Case 717
            vla_step_text = "Put gtext-found into cell L11. [line 1216]" ' vla:4407
            Exit Function
        Case 718
            vla_step_text = "Set gtext-found to how many cells in range Z1:AA3 contain ""banana"". [line 1217]" ' vla:4408
            Exit Function
        Case 719
            vla_step_text = "Put gtext-found into cell L12. [line 1218]" ' vla:4409
            Exit Function
        Case 720
            vla_step_text = "Set gtext-found to how many cells in column Z contain ""zzz"". [line 1219]" ' vla:4410
            Exit Function
        Case 721
            vla_step_text = "Put gtext-found into cell L13. [line 1220]" ' vla:4411
            Exit Function
        Case 722
            vla_step_text = "Set gtext-found to the row of the first cell in range Z1:AA3 containing ""kiwi"". [line 1221]" ' vla:4412
            Exit Function
        Case 723
            vla_step_text = "Put gtext-found into cell L14. [line 1222]" ' vla:4413
            Exit Function
        Case 724
            vla_step_text = "If range Z1:AA3 contains ""split"", put ""yes"" into cell L15. [line 1223]" ' vla:4414
            Exit Function
        Case 725
            vla_step_text = "If column Z does not contain ""kiwi"", put ""no kiwi"" into cell L16. [line 1224]" ' vla:4415
            Exit Function
        Case 726
            vla_step_text = "If cell AA1 contains ""NAN"", put ""contains"" into cell L17. [line 1238]" ' vla:4416
            Exit Function
        Case 727
            vla_step_text = "If cell AA1 starts with ""ban"", put ""starts"" into cell L18. [line 1239]" ' vla:4417
            Exit Function
        Case 728
            vla_step_text = "Put ""kept"" into cell L19. [line 1240]" ' vla:4418
            Exit Function
        Case 729
            vla_step_text = "If cell AA1 does not contain ""nan"", put ""wrong"" into cell L19. [line 1241]" ' vla:4419
            Exit Function
        Case 730
            vla_step_text = "If cell AA1 does not contain ""kiwi"", put ""lacks"" into cell L20. [line 1242]" ' vla:4420
            Exit Function
        Case 731
            vla_step_text = "Put ""kept"" into cell L21. [line 1243]" ' vla:4421
            Exit Function
        Case 732
            vla_step_text = "If cell AA1 contains ""kiwi"", put ""wrong"" into cell L21. [line 1244]" ' vla:4422
            Exit Function
        Case 733
            vla_step_text = "Put ""kept"" into cell L22. [line 1245]" ' vla:4423
            Exit Function
        Case 734
            vla_step_text = "If cell AA1 starts with ""nan"", put ""wrong"" into cell L22. [line 1246]" ' vla:4424
            Exit Function
        Case 735
            vla_step_text = "Check-text-conditions. [line 1248]" ' vla:4425
            Exit Function
        Case 736
            vla_step_text = "Work on sheet GFormula. [line 1259]" ' vla:4426
            Exit Function
        Case 737
            vla_step_text = "Clear everything from A1:H10. [line 1260]" ' vla:4427
            Exit Function
        Case 738
            vla_step_text = "Put 2 into cell B2. [line 1261]" ' vla:4428
            Exit Function
        Case 739
            vla_step_text = "Put 3 into cell B3. [line 1262]" ' vla:4429
            Exit Function
        Case 740
            vla_step_text = "Put 4 into cell B4. [line 1263]" ' vla:4430
            Exit Function
        Case 741
            vla_step_text = "Put 10 into cell C2. [line 1264]" ' vla:4431
            Exit Function
        Case 742
            vla_step_text = "Put 20 into cell C3. [line 1265]" ' vla:4432
            Exit Function
        Case 743
            vla_step_text = "Put 30 into cell C4. [line 1266]" ' vla:4433
            Exit Function
        Case 744
            vla_step_text = "Put formula ""=B2*C2"" into range D2:D4. [line 1267]" ' vla:4434
            Exit Function
        Case 745
            vla_step_text = "Set gf-last to last filled row of column B. [line 1268]" ' vla:4435
            Exit Function
        Case 746
            vla_step_text = "Put formula ""=B2+C2"" into rows 2 to gf-last of column E. [line 1269]" ' vla:4436
            Exit Function
        Case 747
            vla_step_text = "Put ""head"" into cell F1. [line 1270]" ' vla:4437
            Exit Function
        Case 748
            vla_step_text = "Put formula ""=B2-C2"" into rows 2 through 1 of column F. [line 1271]" ' vla:4438
            Exit Function
        Case 749
            vla_step_text = "Set gf-top to largest of range D2:D4. [line 1272]" ' vla:4439
            Exit Function
        Case 750
            vla_step_text = "Put gf-top into cell H1. [line 1273]" ' vla:4440
            Exit Function
        Case 751
            vla_step_text = "Set gf-bottom to smallest of range D2:D4. [line 1274]" ' vla:4441
            Exit Function
        Case 752
            vla_step_text = "Put gf-bottom into cell H2. [line 1275]" ' vla:4442
            Exit Function
        Case 753
            vla_step_text = "Put average of range B2:B4 into cell H3. [line 1276]" ' vla:4443
            Exit Function
        Case 754
            vla_step_text = "Put largest of range C2:C4 into cell H4. [line 1277]" ' vla:4444
            Exit Function
        Case 755
            vla_step_text = "Put smallest of range C2:C4 in cell H5. [line 1278]" ' vla:4445
            Exit Function
        Case 756
            vla_step_text = "Put formula ""="""""""""" into cell B6. [line 1279]" ' vla:4446
            Exit Function
        Case 757
            vla_step_text = "Set gf-empty to count of empty cells in range B2:B7. [line 1280]" ' vla:4447
            Exit Function
        Case 758
            vla_step_text = "Put gf-empty into cell H6. [line 1281]" ' vla:4448
            Exit Function
        Case 759
            vla_step_text = "Set gf-filled to count of filled cells in range B2:B7. [line 1282]" ' vla:4449
            Exit Function
        Case 760
            vla_step_text = "Put gf-filled into cell H7. [line 1283]" ' vla:4450
            Exit Function
        Case 761
            vla_step_text = "Check-formulas. [line 1285]" ' vla:4451
            Exit Function
        Case 762
            vla_step_text = "Work on sheet GStats. [line 1300]" ' vla:4452
            Exit Function
        Case 763
            vla_step_text = "Clear everything from A1:H10. [line 1301]" ' vla:4453
            Exit Function
        Case 764
            vla_step_text = "Put 2 into cell B2. [line 1302]" ' vla:4454
            Exit Function
        Case 765
            vla_step_text = "Put 4 into cell B3. [line 1303]" ' vla:4455
            Exit Function
        Case 766
            vla_step_text = "Put 6 into cell B4. [line 1304]" ' vla:4456
            Exit Function
        Case 767
            vla_step_text = "Put 6 into cell B5. [line 1305]" ' vla:4457
            Exit Function
        Case 768
            vla_step_text = "Put 7 into cell B6. [line 1306]" ' vla:4458
            Exit Function
        Case 769
            vla_step_text = "Set gs-middle to median of range B2:B6. [line 1307]" ' vla:4459
            Exit Function
        Case 770
            vla_step_text = "Put gs-middle into cell H1. [line 1308]" ' vla:4460
            Exit Function
        Case 771
            vla_step_text = "Set gs-spread to standard deviation of range B2:B6 as a sample. [line 1309]" ' vla:4461
            Exit Function
        Case 772
            vla_step_text = "Put gs-spread into cell H2. [line 1310]" ' vla:4462
            Exit Function
        Case 773
            vla_step_text = "Put median of range B2:B5 into cell H3. [line 1311]" ' vla:4463
            Exit Function
        Case 774
            vla_step_text = "Put standard deviation of range B2:B6 as the population into cell H4. [line 1312]" ' vla:4464
            Exit Function
        Case 775
            vla_step_text = "Put ""West"" into cell D2. [line 1313]" ' vla:4465
            Exit Function
        Case 776
            vla_step_text = "Put ""East"" into cell D3. [line 1314]" ' vla:4466
            Exit Function
        Case 777
            vla_step_text = "Put ""West"" into cell D4. [line 1315]" ' vla:4467
            Exit Function
        Case 778
            vla_step_text = "Put ""West"" into cell D5. [line 1316]" ' vla:4468
            Exit Function
        Case 779
            vla_step_text = "Put 10 into cell E2. [line 1317]" ' vla:4469
            Exit Function
        Case 780
            vla_step_text = "Put 20 into cell E3. [line 1318]" ' vla:4470
            Exit Function
        Case 781
            vla_step_text = "Put 30 into cell E5. [line 1319]" ' vla:4471
            Exit Function
        Case 782
            vla_step_text = "Set gs-west to average of range E2:E5 where range D2:D5 matches ""West"". [line 1320]" ' vla:4472
            Exit Function
        Case 783
            vla_step_text = "Put gs-west into cell H5. [line 1321]" ' vla:4473
            Exit Function
        Case 784
            vla_step_text = "Put ""kept"" into cell H6. [line 1322]" ' vla:4474
            Exit Function
        Case 785
            vla_step_text = "Try: [line 1323]" ' vla:4475
            Exit Function
        Case 786
            vla_step_text = "Set gs-none to median of range G1:G3. [line 1324]" ' vla:4476
            Exit Function
        Case 787
            vla_step_text = "Put ""no stop"" into cell H6. [line 1325]" ' vla:4477
            Exit Function
        Case 788
            vla_step_text = "Put the problem into cell H7. [line 1328]" ' vla:4478
            Exit Function
        Case 789
            vla_step_text = "Check-statistics. [line 1330]" ' vla:4479
            Exit Function
        Case 790
            vla_step_text = "Work on sheet GCalc. [line 1343]" ' vla:4480
            Exit Function
        Case 791
            vla_step_text = "Clear everything from A1:D10. [line 1344]" ' vla:4481
            Exit Function
        Case 792
            vla_step_text = "Put 2 into cell A1. [line 1345]" ' vla:4482
            Exit Function
        Case 793
            vla_step_text = "Put formula ""=A1*10"" into cell B1. [line 1346]" ' vla:4483
            Exit Function
        Case 794
            vla_step_text = "Turn off automatic calculation. [line 1347]" ' vla:4484
            Exit Function
        Case 795
            vla_step_text = "Put 3 into cell A1. [line 1348]" ' vla:4485
            Exit Function
        Case 796
            vla_step_text = "Set gc-stale to cell B1. [line 1349]" ' vla:4486
            Exit Function
        Case 797
            vla_step_text = "Put gc-stale into cell C1. [line 1350]" ' vla:4487
            Exit Function
        Case 798
            vla_step_text = "Recalculate this sheet. [line 1351]" ' vla:4488
            Exit Function
        Case 799
            vla_step_text = "Set gc-sheet to cell B1. [line 1352]" ' vla:4489
            Exit Function
        Case 800
            vla_step_text = "Put gc-sheet into cell C2. [line 1353]" ' vla:4490
            Exit Function
        Case 801
            vla_step_text = "Put 4 into cell A1. [line 1354]" ' vla:4491
            Exit Function
        Case 802
            vla_step_text = "Recalculate all open workbooks. [line 1355]" ' vla:4492
            Exit Function
        Case 803
            vla_step_text = "Set gc-all to cell B1. [line 1356]" ' vla:4493
            Exit Function
        Case 804
            vla_step_text = "Put gc-all into cell C3. [line 1357]" ' vla:4494
            Exit Function
        Case 805
            vla_step_text = "Put 5 into cell A1. [line 1358]" ' vla:4495
            Exit Function
        Case 806
            vla_step_text = "Turn on automatic calculation. [line 1359]" ' vla:4496
            Exit Function
        Case 807
            vla_step_text = "Set gc-on to cell B1. [line 1360]" ' vla:4497
            Exit Function
        Case 808
            vla_step_text = "Put gc-on into cell C4. [line 1361]" ' vla:4498
            Exit Function
        Case 809
            vla_step_text = "Put 6 into cell A1. [line 1362]" ' vla:4499
            Exit Function
        Case 810
            vla_step_text = "Set gc-auto to cell B1. [line 1363]" ' vla:4500
            Exit Function
        Case 811
            vla_step_text = "Put gc-auto into cell C5. [line 1364]" ' vla:4501
            Exit Function
        Case 812
            vla_step_text = "Check-calculation. [line 1366]" ' vla:4502
            Exit Function
        Case 813
            vla_step_text = "Work on sheet GFigures. [line 1380]" ' vla:4503
            Exit Function
        Case 814
            vla_step_text = "Clear everything from A1:J10. [line 1381]" ' vla:4504
            Exit Function
        Case 815
            vla_step_text = "Put 2 into cell B2. [line 1382]" ' vla:4505
            Exit Function
        Case 816
            vla_step_text = "Put 4 into cell B3. [line 1383]" ' vla:4506
            Exit Function
        Case 817
            vla_step_text = "Put 6 into cell B4. [line 1384]" ' vla:4507
            Exit Function
        Case 818
            vla_step_text = "Put 6 into cell B5. [line 1385]" ' vla:4508
            Exit Function
        Case 819
            vla_step_text = "Put 7 into cell B6. [line 1386]" ' vla:4509
            Exit Function
        Case 820
            vla_step_text = "Remember range B2:B6 as gfig-values. [line 1387]" ' vla:4510
            Exit Function
        Case 821
            vla_step_text = "Set gfig-middle to median of gfig-values. [line 1388]" ' vla:4511
            Exit Function
        Case 822
            vla_step_text = "Put gfig-middle into cell H1. [line 1389]" ' vla:4512
            Exit Function
        Case 823
            vla_step_text = "Put median of gfig-values into cell H2. [line 1390]" ' vla:4513
            Exit Function
        Case 824
            vla_step_text = "Set gfig-spread to standard deviation of gfig-values as a sample. [line 1391]" ' vla:4514
            Exit Function
        Case 825
            vla_step_text = "Put gfig-spread into cell H3. [line 1392]" ' vla:4515
            Exit Function
        Case 826
            vla_step_text = "Put standard deviation of gfig-values as the population into cell H4. [line 1393]" ' vla:4516
            Exit Function
        Case 827
            vla_step_text = "Put ""West"" into cell D2. [line 1394]" ' vla:4517
            Exit Function
        Case 828
            vla_step_text = "Put ""East"" into cell D3. [line 1395]" ' vla:4518
            Exit Function
        Case 829
            vla_step_text = "Put ""West"" into cell D4. [line 1396]" ' vla:4519
            Exit Function
        Case 830
            vla_step_text = "Put ""West"" into cell D5. [line 1397]" ' vla:4520
            Exit Function
        Case 831
            vla_step_text = "Put 10 into cell E2. [line 1398]" ' vla:4521
            Exit Function
        Case 832
            vla_step_text = "Put 20 into cell E3. [line 1399]" ' vla:4522
            Exit Function
        Case 833
            vla_step_text = "Put 30 into cell E5. [line 1400]" ' vla:4523
            Exit Function
        Case 834
            vla_step_text = "Put sum of range E2:E5 where range D2:D5 matches ""West"" into cell H5. [line 1401]" ' vla:4524
            Exit Function
        Case 835
            vla_step_text = "Put average of range E2:E5 where range D2:D5 matches ""West"" in cell H6. [line 1402]" ' vla:4525
            Exit Function
        Case 836
            vla_step_text = "Format cell C2 as dollars. [line 1403]" ' vla:4526
            Exit Function
        Case 837
            vla_step_text = "Put formula ""=B2*10"" into cell C2 and fill down to the last filled row of column B. [line 1404]" ' vla:4527
            Exit Function
        Case 838
            vla_step_text = "Put ""head"" into cell F1. [line 1405]" ' vla:4528
            Exit Function
        Case 839
            vla_step_text = "Put 1 into cell G2. [line 1406]" ' vla:4529
            Exit Function
        Case 840
            vla_step_text = "Put formula ""=B2*10"" into cell F2 and fill down to the last filled row of column G. [line 1407]" ' vla:4530
            Exit Function
        Case 841
            vla_step_text = "Put formula ""=B2*10"" in cell J2 and fill down to the last filled row of column I. [line 1408]" ' vla:4531
            Exit Function
        Case 842
            vla_step_text = "Check-figures. [line 1410]" ' vla:4532
            Exit Function
        Case 843
            vla_step_text = "Work on sheet LPhrases. [line 1422]" ' vla:4533
            Exit Function
        Case 844
            vla_step_text = "Clear everything from A1:H10. [line 1423]" ' vla:4534
            Exit Function
        Case 845
            vla_step_text = "Put 2 into cell B2. [line 1424]" ' vla:4535
            Exit Function
        Case 846
            vla_step_text = "Put 4 into cell B3. [line 1425]" ' vla:4536
            Exit Function
        Case 847
            vla_step_text = "Put 6 into cell B4. [line 1426]" ' vla:4537
            Exit Function
        Case 848
            vla_step_text = "Put 6 into cell B5. [line 1427]" ' vla:4538
            Exit Function
        Case 849
            vla_step_text = "Put 7 into cell B6. [line 1428]" ' vla:4539
            Exit Function
        Case 850
            vla_step_text = "Remember range B2:B6 as lph-values. [line 1429]" ' vla:4540
            Exit Function
        Case 851
            vla_step_text = "If median of lph-values is more than 5, put ""above"" into cell H1. [line 1430]" ' vla:4541
            Exit Function
        Case 852
            vla_step_text = "Put median of lph-values plus 1 into cell H2. [line 1431]" ' vla:4542
            Exit Function
        Case 853
            vla_step_text = "Put standard deviation of range B2:B6 as a sample into cell H3. [line 1432]" ' vla:4543
            Exit Function
        Case 854
            vla_step_text = "Set lph-gap to largest of range B2:B6 minus smallest of range B2:B6. [line 1433]" ' vla:4544
            Exit Function
        Case 855
            vla_step_text = "Put lph-gap into cell H4. [line 1434]" ' vla:4545
            Exit Function
        Case 856
            vla_step_text = "Put last filled row of column B plus 1 into cell H5. [line 1435]" ' vla:4546
            Exit Function
        Case 857
            vla_step_text = "If standard deviation of lph-values as the population is less than 2, put ""narrow"" into cell H6. [line 1436]" ' vla:4547
            Exit Function
        Case 858
            vla_step_text = "Check-phrases. [line 1438]" ' vla:4548
            Exit Function
        Case 859
            vla_step_text = "Go to sheet Output. [line 1440]" ' vla:4549
            Exit Function
        Case 860
            vla_step_text = "Tidy-up. [line 1442]" ' vla:4550
            Exit Function
        Case 861
            vla_step_text = "Turn on screen updating. [line 1443]" ' vla:4551
            Exit Function
        Case 862
            vla_step_text = "Log ""report finished"". [line 1444]" ' vla:4552
            Exit Function
        Case Else
            vla_step_text = "an unknown step" ' vla:4553
            Exit Function
    End Select
End Function


