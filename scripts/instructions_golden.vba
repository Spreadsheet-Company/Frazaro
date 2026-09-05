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

Public Sub main()
    Dim counter As Variant ' vla:71
    Dim grand As Variant ' vla:72
    Dim results As Variant ' vla:73
    Dim r As Variant ' vla:74
    Dim biggest As Variant ' vla:75
    Dim probe As Variant ' vla:76
    Dim round_check As Variant ' vla:77
    Dim thousand_check As Variant ' vla:78
    Dim f_last As Variant ' vla:79
    Dim echo_row As Variant ' vla:80
    Dim check_row As Variant ' vla:81
    Dim stripe_row As Variant ' vla:82
    Dim back_row As Variant ' vla:83
    Dim search_row As Variant ' vla:84
    Dim f As Variant ' vla:85
    Dim list_count As Variant ' vla:86
    Dim region As Variant ' vla:87
    Dim fuel As Variant ' vla:88
    Dim fee As Variant ' vla:89
    Dim full_commission As Variant ' vla:90
    Dim default_commission As Variant ' vla:91
    Dim pick_check As Variant ' vla:92
    Dim ax_price As Variant ' vla:93
    Dim key_count As Variant ' vla:94
    Dim price_sum As Variant ' vla:95
    Dim k As Variant ' vla:96
    Dim pair As Variant ' vla:97
    Dim qty_values As Variant ' vla:98
    Dim q As Variant ' vla:99
    On Error GoTo vla_fail ' vla:100
    vla_step = 1 ' vla:101
    If vlatraceon() Then ' vla:102
        Call vlatracestep(1, vla_step_text(1)) ' vla:102
    End If
    ' ---- instructions.txt:10 Work on sheet Output. ----
    Call vlaensuresheet("output") ' vla:105 src:18
    Call worksheets("output").activate
    vla_step = 2 ' vla:106
    If vlatraceon() Then ' vla:107
        Call vlatracestep(2, vla_step_text(2)) ' vla:107
    End If
    ' ---- instructions.txt:20 Turn off screen updating. ----
    application.screenupdating = False
    vla_step = 3 ' vla:111
    If vlatraceon() Then ' vla:112
        Call vlatracestep(3, vla_step_text(3)) ' vla:112
    End If
    Dim total As Double ' vla:114 src:21
    vla_step = 4 ' vla:115
    If vlatraceon() Then ' vla:116
        Call vlatracestep(4, vla_step_text(4)) ' vla:116
    End If
    total = 0 ' vla:118 src:22
    vla_step = 5 ' vla:119
    If vlatraceon() Then ' vla:120
        Call vlatracestep(5, vla_step_text(5)) ' vla:120
    End If
    range("a1") = "Test Report" ' vla:122 src:23
    vla_step = 6 ' vla:123
    If vlatraceon() Then ' vla:124
        Call vlatracestep(6, vla_step_text(6)) ' vla:124
    End If
    range("a1").font.bold = True
    vla_step = 7 ' vla:127
    If vlatraceon() Then ' vla:128
        Call vlatracestep(7, vla_step_text(7)) ' vla:128
    End If
    range("a1").font.size = 14
    vla_step = 8 ' vla:131
    If vlatraceon() Then ' vla:132
        Call vlatracestep(8, vla_step_text(8)) ' vla:132
    End If
    range("d1") = date() ' vla:134 src:26
    vla_step = 9 ' vla:135
    If vlatraceon() Then ' vla:136
        Call vlatracestep(9, vla_step_text(9)) ' vla:136
    End If
    ' ---- instructions.txt:28 Repeat 5 times: ----
    For counter = 1 To 5
        vla_step = 10 ' vla:140 src:28
        If vlatraceon() Then ' vla:141 src:28
            Call vlatracestep(10, vla_step_text(10)) ' vla:141 src:28
        End If
        total = (total + counter)
        vla_step = 11 ' vla:144 src:28
        If vlatraceon() Then ' vla:145 src:28
            Call vlatracestep(11, vla_step_text(11)) ' vla:145 src:28
        End If
        If ((counter Mod 2) = 0) Then ' vla:147 src:30
            Debug.Print ("even step " & counter) ' vla:147 src:30
        End If
    Next counter
    vla_step = 12 ' vla:148
    If vlatraceon() Then ' vla:149
        Call vlatracestep(12, vla_step_text(12)) ' vla:149
    End If
    ' ---- instructions.txt:32 Put total into cell B2. ----
    range("b2") = total ' vla:152 src:32
    vla_step = 13 ' vla:153
    If vlatraceon() Then ' vla:154
        Call vlatracestep(13, vla_step_text(13)) ' vla:154
    End If
    range("b3").Formula2 = "=B2*2"
    vla_step = 14 ' vla:157
    If vlatraceon() Then ' vla:158
        Call vlatracestep(14, vla_step_text(14)) ' vla:158
    End If
    range("b4") = application.worksheetfunction.sum(range("B2:B3")) ' vla:160 src:37
    vla_step = 15 ' vla:161
    If vlatraceon() Then ' vla:162
        Call vlatracestep(15, vla_step_text(15)) ' vla:162
    End If
    grand = application.worksheetfunction.sum(range("b2:b3")) ' vla:164 src:38
    vla_step = 16 ' vla:165
    If vlatraceon() Then ' vla:166
        Call vlatracestep(16, vla_step_text(16)) ' vla:166
    End If
    Debug.Print ("grand is " & grand) ' vla:168 src:39
    vla_step = 17 ' vla:169
    If vlatraceon() Then ' vla:170
        Call vlatracestep(17, vla_step_text(17)) ' vla:170
    End If
    ' ---- instructions.txt:41 If grand is greater than 40: ----
    If (grand > 40) Then ' vla:173 src:43
        vla_step = 18 ' vla:175 src:43
        If vlatraceon() Then ' vla:176 src:43
            Call vlatracestep(18, vla_step_text(18)) ' vla:176 src:43
        End If
        range("c4") = "PASS" ' vla:178 src:44
        vla_step = 19 ' vla:179 src:43
        If vlatraceon() Then ' vla:180 src:43
            Call vlatracestep(19, vla_step_text(19)) ' vla:180 src:43
        End If
        For counter = 1 To 2
            vla_step = 20 ' vla:183 src:45
            If vlatraceon() Then ' vla:184 src:45
                Call vlatracestep(20, vla_step_text(20)) ' vla:184 src:45
            End If
            Debug.Print ("pass check " & counter) ' vla:186 src:46
        Next counter
        vla_step = 21 ' vla:187 src:43
        If vlatraceon() Then ' vla:188 src:43
            Call vlatracestep(21, vla_step_text(21)) ' vla:188 src:43
        End If
        range("c4").font.bold = True
    Else
        vla_step = 22 ' vla:192 src:43
        If vlatraceon() Then ' vla:193 src:43
            Call vlatracestep(22, vla_step_text(22)) ' vla:193 src:43
        End If
        ' ---- instructions.txt:50 Otherwise: ----
        range("c4") = "CHECK" ' vla:196 src:51
    End If
    vla_step = 23 ' vla:197
    If vlatraceon() Then ' vla:198
        Call vlatracestep(23, vla_step_text(23)) ' vla:198
    End If
    ' ---- instructions.txt:53 If grand is at least 45, make cell C4 yellow. ----
    If (grand >= 45) Then ' vla:201 src:53
        range("c4").interior.color = vbyellow
    End If
    vla_step = 24 ' vla:202
    If vlatraceon() Then ' vla:203
        Call vlatracestep(24, vla_step_text(24)) ' vla:203
    End If
    Dim label As String ' vla:205 src:54
    vla_step = 25 ' vla:206
    If vlatraceon() Then ' vla:207
        Call vlatracestep(25, vla_step_text(25)) ' vla:207
    End If
    label = ("Total: " & total) ' vla:209 src:55
    vla_step = 26 ' vla:210
    If vlatraceon() Then ' vla:211
        Call vlatracestep(26, vla_step_text(26)) ' vla:211
    End If
    range("a6") = label ' vla:213 src:56
    vla_step = 27 ' vla:214
    If vlatraceon() Then ' vla:215
        Call vlatracestep(27, vla_step_text(27)) ' vla:215
    End If
    range("a6").font.color = vlacolor(hot_pink)
    vla_step = 28 ' vla:218
    If vlatraceon() Then ' vla:219
        Call vlatracestep(28, vla_step_text(28)) ' vla:219
    End If
    Set results = range("b2:b4") ' vla:221 src:58
    vla_step = 29 ' vla:222
    If vlatraceon() Then ' vla:223
        Call vlatracestep(29, vla_step_text(29)) ' vla:223
    End If
    For Each r In results ' vla:225 src:59
        Debug.Print r ' vla:225 src:59
    Next r
    vla_step = 30 ' vla:226
    If vlatraceon() Then ' vla:227
        Call vlatracestep(30, vla_step_text(30)) ' vla:227
    End If
    biggest = application.worksheetfunction.max(results) ' vla:229 src:60
    vla_step = 31 ' vla:230
    If vlatraceon() Then ' vla:231
        Call vlatracestep(31, vla_step_text(31)) ' vla:231
    End If
    Debug.Print ((("largest result is " & biggest) & ", label length ") & len(label)) ' vla:233 src:61
    vla_step = 32 ' vla:234
    If vlatraceon() Then ' vla:235
        Call vlatracestep(32, vla_step_text(32)) ' vla:235
    End If
    ' ---- instructions.txt:63 Repeat 3 times: ----
    For counter = 1 To 3
        vla_step = 33 ' vla:239 src:64
        If vlatraceon() Then ' vla:240 src:64
            Call vlatracestep(33, vla_step_text(33)) ' vla:240 src:64
        End If
        cells(counter, "e") = (counter * 10) ' vla:242 src:65
    Next counter
    vla_step = 34 ' vla:243
    If vlatraceon() Then ' vla:244
        Call vlatracestep(34, vla_step_text(34)) ' vla:244
    End If
    ' ---- instructions.txt:67 Set probe to cell in column E row 2. ----
    probe = cells(2, "e") ' vla:247 src:67
    vla_step = 35 ' vla:248
    If vlatraceon() Then ' vla:249
        Call vlatracestep(35, vla_step_text(35)) ' vla:249
    End If
    Dim num_col_check As String ' vla:251 src:68
    vla_step = 36 ' vla:252
    If vlatraceon() Then ' vla:253
        Call vlatracestep(36, vla_step_text(36)) ' vla:253
    End If
    If (cells(2, 5) = 20) Then ' vla:255 src:69
        num_col_check = "yes" ' vla:255 src:69
    End If
    vla_step = 37 ' vla:256
    If vlatraceon() Then ' vla:257
        Call vlatracestep(37, vla_step_text(37)) ' vla:257
    End If
    Dim value_word_check As String ' vla:259 src:70
    vla_step = 38 ' vla:260
    If vlatraceon() Then ' vla:261
        Call vlatracestep(38, vla_step_text(38)) ' vla:261
    End If
    If (cells(3, 5) = 30) Then ' vla:263 src:71
        value_word_check = "ok" ' vla:263 src:71
    End If
    vla_step = 39 ' vla:264
    If vlatraceon() Then ' vla:265
        Call vlatracestep(39, vla_step_text(39)) ' vla:265
    End If
    range("output!h16") = "bang" ' vla:267 src:72
    vla_step = 40 ' vla:268
    If vlatraceon() Then ' vla:269
        Call vlatracestep(40, vla_step_text(40)) ' vla:269
    End If
    range("a1").horizontalalignment = xlcenter
    vla_step = 41 ' vla:272
    If vlatraceon() Then ' vla:273
        Call vlatracestep(41, vla_step_text(41)) ' vla:273
    End If
    columns("d").columnwidth = 24
    vla_step = 42 ' vla:276
    If vlatraceon() Then ' vla:277
        Call vlatracestep(42, vla_step_text(42)) ' vla:277
    End If
    round_check = round(3.14159, 2) ' vla:279 src:75
    vla_step = 43 ' vla:280
    If vlatraceon() Then ' vla:281
        Call vlatracestep(43, vla_step_text(43)) ' vla:281
    End If
    Debug.Print ("rounded is " & round_check) ' vla:283 src:76
    vla_step = 44 ' vla:284
    If vlatraceon() Then ' vla:285
        Call vlatracestep(44, vla_step_text(44)) ' vla:285
    End If
    thousand_check = (1000 + 500) ' vla:287 src:77
    vla_step = 45 ' vla:288
    If vlatraceon() Then ' vla:289
        Call vlatracestep(45, vla_step_text(45)) ' vla:289
    End If
    Debug.Print ("thousands read as " & thousand_check) ' vla:291 src:78
    vla_step = 46 ' vla:292
    If vlatraceon() Then ' vla:293
        Call vlatracestep(46, vla_step_text(46)) ' vla:293
    End If
    Debug.Print ("probe is " & probe) ' vla:295 src:79
    vla_step = 47 ' vla:296
    If vlatraceon() Then ' vla:297
        Call vlatracestep(47, vla_step_text(47)) ' vla:297
    End If
    ' ---- instructions.txt:81 Create a number called countdown. ----
    Dim countdown As Double ' vla:300 src:81
    vla_step = 48 ' vla:301
    If vlatraceon() Then ' vla:302
        Call vlatracestep(48, vla_step_text(48)) ' vla:302
    End If
    countdown = 3 ' vla:304 src:82
    vla_step = 49 ' vla:305
    If vlatraceon() Then ' vla:306
        Call vlatracestep(49, vla_step_text(49)) ' vla:306
    End If
    Do While (countdown > 0) ' vla:308 src:83
        vla_step = 50 ' vla:309 src:83
        If vlatraceon() Then ' vla:310 src:83
            Call vlatracestep(50, vla_step_text(50)) ' vla:310 src:83
        End If
        Debug.Print ("countdown " & countdown) ' vla:312 src:84
        vla_step = 51 ' vla:313 src:83
        If vlatraceon() Then ' vla:314 src:83
            Call vlatracestep(51, vla_step_text(51)) ' vla:314 src:83
        End If
        countdown = (countdown - 1) ' vla:316 src:85
    Loop
    vla_step = 53 ' vla:317
    If vlatraceon() Then ' vla:318
        Call vlatracestep(53, vla_step_text(53)) ' vla:318
    End If
    ' ---- instructions.txt:92 Stamp. ----
    Call stamp ' vla:321 src:92
    vla_step = 54 ' vla:322
    If vlatraceon() Then ' vla:323
        Call vlatracestep(54, vla_step_text(54)) ' vla:323
    End If
    Call stamp(row_number:=2, value:="beta") ' vla:325 src:93
    vla_step = 55 ' vla:326
    If vlatraceon() Then ' vla:327
        Call vlatracestep(55, vla_step_text(55)) ' vla:327
    End If
    f_last = cells(rows.count, "f").end(xlup).row ' vla:329 src:94
    vla_step = 56 ' vla:330
    If vlatraceon() Then ' vla:331
        Call vlatracestep(56, vla_step_text(56)) ' vla:331
    End If
    Debug.Print ("column F filled to row " & f_last) ' vla:333 src:95
    vla_step = 57 ' vla:334
    If vlatraceon() Then ' vla:335
        Call vlatracestep(57, vla_step_text(57)) ' vla:335
    End If
    For echo_row = 1 To f_last ' vla:337 src:96
        Debug.Print ("echo " & echo_row) ' vla:337 src:96
    Next echo_row
    vla_step = 58 ' vla:338
    If vlatraceon() Then ' vla:339
        Call vlatracestep(58, vla_step_text(58)) ' vla:339
    End If
    For check_row = 1 To f_last ' vla:341 src:97
        vla_step = 59 ' vla:342 src:97
        If vlatraceon() Then ' vla:343 src:97
            Call vlatracestep(59, vla_step_text(59)) ' vla:343 src:97
        End If
        If (instr(1, cells(check_row, "f"), "ok", vbtextcompare) > 0) Then ' vla:345 src:98
            cells(check_row, "f").font.bold = True
        End If
    Next check_row
    vla_step = 60 ' vla:346
    If vlatraceon() Then ' vla:347
        Call vlatracestep(60, vla_step_text(60)) ' vla:347
    End If
    ' ---- instructions.txt:100 Count stripe-row from 1 to f-last step 2: ----
    For stripe_row = 1 To f_last Step 2 ' vla:350 src:103
        vla_step = 61 ' vla:351 src:103
        If vlatraceon() Then ' vla:352 src:103
            Call vlatracestep(61, vla_step_text(61)) ' vla:352 src:103
        End If
        cells(stripe_row, "f").font.bold = True
    Next stripe_row
    vla_step = 62 ' vla:355
    If vlatraceon() Then ' vla:356
        Call vlatracestep(62, vla_step_text(62)) ' vla:356
    End If
    ' ---- instructions.txt:106 Count back-row down from f-last to 1 step 2: ----
    For back_row = f_last To 1 Step (0 - 2) ' vla:359 src:106
        vla_step = 63 ' vla:360 src:106
        If vlatraceon() Then ' vla:361 src:106
            Call vlatracestep(63, vla_step_text(63)) ' vla:361 src:106
        End If
        Debug.Print ("back-row " & back_row) ' vla:363 src:107
    Next back_row
    vla_step = 64 ' vla:364
    If vlatraceon() Then ' vla:365
        Call vlatracestep(64, vla_step_text(64)) ' vla:365
    End If
    ' ---- instructions.txt:109 Count search-row from 1 to 10: ----
    For search_row = 1 To 10 ' vla:368 src:111
        vla_step = 65 ' vla:369 src:111
        If vlatraceon() Then ' vla:370 src:111
            Call vlatracestep(65, vla_step_text(65)) ' vla:370 src:111
        End If
        If (search_row = 3) Then ' vla:372 src:112
            Exit For ' vla:372 src:112
        End If
    Next search_row
    vla_step = 66 ' vla:373
    If vlatraceon() Then ' vla:374
        Call vlatracestep(66, vla_step_text(66)) ' vla:374
    End If
    ' ---- instructions.txt:114 Log "stopped at " joined with search-row. ----
    Debug.Print ("stopped at " & search_row) ' vla:377 src:114
    vla_step = 67 ' vla:378
    If vlatraceon() Then ' vla:379
        Call vlatracestep(67, vla_step_text(67)) ' vla:379
    End If
    ' ---- instructions.txt:116 Create a list called found-items. ----
    Dim found_items As Collection ' vla:382 src:119
    Set found_items = New Collection ' vla:382 src:119
    vla_step = 68 ' vla:383
    If vlatraceon() Then ' vla:384
        Call vlatracestep(68, vla_step_text(68)) ' vla:384
    End If
    For counter = 1 To 4
        vla_step = 69 ' vla:387 src:120
        If vlatraceon() Then ' vla:388 src:120
            Call vlatracestep(69, vla_step_text(69)) ' vla:388 src:120
        End If
        If (counter > 2) Then ' vla:390 src:121
            Call found_items.add((counter * 100)) ' vla:390 src:121
        End If
    Next counter
    vla_step = 70 ' vla:391
    If vlatraceon() Then ' vla:392
        Call vlatracestep(70, vla_step_text(70)) ' vla:392
    End If
    ' ---- instructions.txt:123 For each f in found-items, log "found " joined with f. ----
    For Each f In found_items ' vla:395 src:123
        Debug.Print ("found " & f) ' vla:395 src:123
    Next f
    vla_step = 71 ' vla:396
    If vlatraceon() Then ' vla:397
        Call vlatracestep(71, vla_step_text(71)) ' vla:397
    End If
    list_count = vlacount(found_items) ' vla:399 src:124
    vla_step = 72 ' vla:400
    If vlatraceon() Then ' vla:401
        Call vlatracestep(72, vla_step_text(72)) ' vla:401
    End If
    Debug.Print ("list holds " & list_count) ' vla:403 src:125
    vla_step = 73 ' vla:404
    If vlatraceon() Then ' vla:405
        Call vlatracestep(73, vla_step_text(73)) ' vla:405
    End If
    ' ---- instructions.txt:127 Create a text called verdict. ----
    Dim verdict As String ' vla:408 src:131
    vla_step = 74 ' vla:409
    If vlatraceon() Then ' vla:410
        Call vlatracestep(74, vla_step_text(74)) ' vla:410
    End If
    If (grand > 100) Then ' vla:412 src:132
        vla_step = 75 ' vla:414 src:132
        If vlatraceon() Then ' vla:415 src:132
            Call vlatracestep(75, vla_step_text(75)) ' vla:415 src:132
        End If
        verdict = "huge" ' vla:417 src:133
    ElseIf (grand > 40) Then
        vla_step = 76 ' vla:419 src:132
        If vlatraceon() Then ' vla:420 src:132
            Call vlatracestep(76, vla_step_text(76)) ' vla:420 src:132
        End If
        ' ---- instructions.txt:135 Otherwise, if grand is greater than 40: ----
        verdict = "solid" ' vla:423 src:136
    Else
        vla_step = 77 ' vla:425 src:132
        If vlatraceon() Then ' vla:426 src:132
            Call vlatracestep(77, vla_step_text(77)) ' vla:426 src:132
        End If
        ' ---- instructions.txt:138 Otherwise: ----
        verdict = "small" ' vla:429 src:139
    End If
    vla_step = 78 ' vla:430
    If vlatraceon() Then ' vla:431
        Call vlatracestep(78, vla_step_text(78)) ' vla:431
    End If
    ' ---- instructions.txt:141 Create a text called region-label. ----
    Dim region_label As String ' vla:434 src:141
    vla_step = 79 ' vla:435
    If vlatraceon() Then ' vla:436
        Call vlatracestep(79, vla_step_text(79)) ' vla:436
    End If
    region = "South" ' vla:438 src:142
    vla_step = 80 ' vla:439
    If vlatraceon() Then ' vla:440
        Call vlatracestep(80, vla_step_text(80)) ' vla:440
    End If
    Select Case region ' vla:442 src:143
        Case "North"
            vla_step = 81 ' vla:444 src:143
            If vlatraceon() Then ' vla:445 src:143
                Call vlatracestep(81, vla_step_text(81)) ' vla:445 src:143
            End If
            region_label = "cold" ' vla:447 src:144
        Case "South", "East"
            vla_step = 82 ' vla:449 src:143
            If vlatraceon() Then ' vla:450 src:143
                Call vlatracestep(82, vla_step_text(82)) ' vla:450 src:143
            End If
            ' ---- instructions.txt:146 When it is "South" or "East": ----
            region_label = "warm" ' vla:453 src:147
        Case Else
            vla_step = 83 ' vla:455 src:143
            If vlatraceon() Then ' vla:456 src:143
                Call vlatracestep(83, vla_step_text(83)) ' vla:456 src:143
            End If
            ' ---- instructions.txt:149 Otherwise: ----
            region_label = "unknown" ' vla:459 src:150
    End Select
    vla_step = 84 ' vla:460
    If vlatraceon() Then ' vla:461
        Call vlatracestep(84, vla_step_text(84)) ' vla:461
    End If
    ' ---- instructions.txt:152 Log "verdict " joined with verdict joined with ", region " joined w... ----
    Debug.Print ((("verdict " & verdict) & ", region ") & region_label) ' vla:464 src:152
    vla_step = 85 ' vla:465
    If vlatraceon() Then ' vla:466
        Call vlatracestep(85, vla_step_text(85)) ' vla:466
    End If
    ' ---- instructions.txt:154 Create a number called until-count. ----
    Dim until_count As Double ' vla:469 src:154
    vla_step = 86 ' vla:470
    If vlatraceon() Then ' vla:471
        Call vlatracestep(86, vla_step_text(86)) ' vla:471
    End If
    fuel = 3 ' vla:473 src:155
    vla_step = 87 ' vla:474
    If vlatraceon() Then ' vla:475
        Call vlatracestep(87, vla_step_text(87)) ' vla:475
    End If
    Do Until (fuel = 0) ' vla:477 src:156
        vla_step = 88 ' vla:478 src:156
        If vlatraceon() Then ' vla:479 src:156
            Call vlatracestep(88, vla_step_text(88)) ' vla:479 src:156
        End If
        fuel = (fuel - 1) ' vla:481 src:157
        vla_step = 89 ' vla:482 src:156
        If vlatraceon() Then ' vla:483 src:156
            Call vlatracestep(89, vla_step_text(89)) ' vla:483 src:156
        End If
        until_count = (until_count + 1)
    Loop
    vla_step = 90 ' vla:486
    If vlatraceon() Then ' vla:487
        Call vlatracestep(90, vla_step_text(90)) ' vla:487
    End If
    ' ---- instructions.txt:160 Log "repeat-until ran " joined with until-count joined with " times". ----
    Debug.Print (("repeat-until ran " & until_count) & " times") ' vla:490 src:160
    vla_step = 91 ' vla:491
    If vlatraceon() Then ' vla:492
        Call vlatracestep(91, vla_step_text(91)) ' vla:492
    End If
    ' ---- instructions.txt:162 Create a text called rescue. ----
    Dim rescue As String ' vla:495 src:166
    vla_step = 92 ' vla:496
    If vlatraceon() Then ' vla:497
        Call vlatracestep(92, vla_step_text(92)) ' vla:497
    End If
    On Error GoTo vla_tryf_1 ' vla:499 src:167
    vla_step = 93 ' vla:500 src:167
    If vlatraceon() Then ' vla:501 src:167
        Call vlatracestep(93, vla_step_text(93)) ' vla:501 src:167
    End If
    Call worksheets("nowhere-land").activate
    vla_step = 94 ' vla:504 src:167
    If vlatraceon() Then ' vla:505 src:167
        Call vlatracestep(94, vla_step_text(94)) ' vla:505 src:167
    End If
    rescue = "unreachable" ' vla:507 src:169
    GoTo vla_tryd_1 ' vla:508 src:167
vla_tryf_1: ' vla:509 src:167
    vla_problem = err.description ' vla:510 src:167
    Resume vla_tryr_1 ' vla:511 src:167
vla_tryr_1: ' vla:512 src:167
    On Error GoTo vla_fail ' vla:513 src:167
    vla_step = 95 ' vla:514 src:167
    If vlatraceon() Then ' vla:515 src:167
        Call vlatracestep(95, vla_step_text(95)) ' vla:515 src:167
    End If
    ' ---- instructions.txt:171 If that fails: ----
    Debug.Print ("the problem was " & vla_problem) ' vla:518 src:172
    vla_step = 96 ' vla:519 src:167
    If vlatraceon() Then ' vla:520 src:167
        Call vlatracestep(96, vla_step_text(96)) ' vla:520 src:167
    End If
    If (Not (len(trim((vla_problem & ""))) = 0)) Then ' vla:522 src:173
        rescue = "rescued" ' vla:522 src:173
    End If
vla_tryd_1: ' vla:523 src:167
    On Error GoTo vla_fail ' vla:524 src:167
    vla_step = 97 ' vla:525
    If vlatraceon() Then ' vla:526
        Call vlatracestep(97, vla_step_text(97)) ' vla:526
    End If
    ' ---- instructions.txt:175 Create a number called risk-free. ----
    Dim risk_free As Double ' vla:529 src:175
    vla_step = 98 ' vla:530
    If vlatraceon() Then ' vla:531
        Call vlatracestep(98, vla_step_text(98)) ' vla:531
    End If
    On Error GoTo vla_tryf_2 ' vla:533 src:176
    vla_step = 99 ' vla:534 src:176
    If vlatraceon() Then ' vla:535 src:176
        Call vlatracestep(99, vla_step_text(99)) ' vla:535 src:176
    End If
    risk_free = 7 ' vla:537 src:177
    GoTo vla_tryd_2 ' vla:538 src:176
vla_tryf_2: ' vla:539 src:176
    vla_problem = err.description ' vla:540 src:176
    Resume vla_tryr_2 ' vla:541 src:176
vla_tryr_2: ' vla:542 src:176
    On Error GoTo vla_fail ' vla:543 src:176
    vla_step = 100 ' vla:544 src:176
    If vlatraceon() Then ' vla:545 src:176
        Call vlatracestep(100, vla_step_text(100)) ' vla:545 src:176
    End If
    ' ---- instructions.txt:179 If that fails: ----
    risk_free = -1 ' vla:548 src:180
vla_tryd_2: ' vla:549 src:176
    On Error GoTo vla_fail ' vla:550 src:176
    vla_step = 101 ' vla:551
    If vlatraceon() Then ' vla:552
        Call vlatracestep(101, vla_step_text(101)) ' vla:552
    End If
    ' ---- instructions.txt:182 Log "rescue " joined with rescue joined with ", risk-free " joined ... ----
    Debug.Print ((("rescue " & rescue) & ", risk-free ") & risk_free) ' vla:555 src:182
    vla_step = 103 ' vla:556
    If vlatraceon() Then ' vla:557
        Call vlatracestep(103, vla_step_text(103)) ' vla:557
    End If
    ' ---- instructions.txt:189 Set fee to tax of 100. ----
    fee = tax(100) ' vla:560 src:189
    vla_step = 104 ' vla:561
    If vlatraceon() Then ' vla:562
        Call vlatracestep(104, vla_step_text(104)) ' vla:562
    End If
    Debug.Print ("fee is " & fee) ' vla:564 src:190
    vla_step = 105 ' vla:565
    If vlatraceon() Then ' vla:566
        Call vlatracestep(105, vla_step_text(105)) ' vla:566
    End If
    Dim fee_size As String ' vla:568 src:191
    vla_step = 106 ' vla:569
    If vlatraceon() Then ' vla:570
        Call vlatracestep(106, vla_step_text(106)) ' vla:570
    End If
    If (tax(50) > 3) Then ' vla:572 src:192
        vla_step = 107 ' vla:574 src:192
        If vlatraceon() Then ' vla:575 src:192
            Call vlatracestep(107, vla_step_text(107)) ' vla:575 src:192
        End If
        fee_size = "big" ' vla:577 src:193
    Else
        vla_step = 108 ' vla:579 src:192
        If vlatraceon() Then ' vla:580 src:192
            Call vlatracestep(108, vla_step_text(108)) ' vla:580 src:192
        End If
        ' ---- instructions.txt:195 Otherwise: ----
        fee_size = "small" ' vla:583 src:196
    End If
    vla_step = 113 ' vla:584
    If vlatraceon() Then ' vla:585
        Call vlatracestep(113, vla_step_text(113)) ' vla:585
    End If
    ' ---- instructions.txt:208 Set full-commission to commission using sale of 2000 and rate of 10%. ----
    full_commission = commission(sale:=2000, rate:=(10 / 100)) ' vla:588 src:208
    vla_step = 114 ' vla:589
    If vlatraceon() Then ' vla:590
        Call vlatracestep(114, vla_step_text(114)) ' vla:590
    End If
    default_commission = commission(sale:=600) ' vla:592 src:209
    vla_step = 115 ' vla:593
    If vlatraceon() Then ' vla:594
        Call vlatracestep(115, vla_step_text(115)) ' vla:594
    End If
    Debug.Print ((("commissions " & full_commission) & " / ") & default_commission) ' vla:596 src:210
    vla_step = 116 ' vla:597
    If vlatraceon() Then ' vla:598
        Call vlatracestep(116, vla_step_text(116)) ' vla:598
    End If
    ' ---- instructions.txt:212 Try: ----
    On Error GoTo vla_tryf_3 ' vla:601 src:219
    vla_step = 117 ' vla:602 src:219
    If vlatraceon() Then ' vla:603 src:219
        Call vlatracestep(117, vla_step_text(117)) ' vla:603 src:219
    End If
    Call vlachecksheetname("Q1 Data")
    Call vlachecksheetabsent("Q1 Data")
    Call worksheets.add
    activesheet.name = "Q1 Data"
    GoTo vla_tryd_3 ' vla:606 src:219
vla_tryf_3: ' vla:607 src:219
    vla_problem = err.description ' vla:608 src:219
    Resume vla_tryr_3 ' vla:609 src:219
vla_tryr_3: ' vla:610 src:219
    On Error GoTo vla_fail ' vla:611 src:219
vla_tryd_3: ' vla:612 src:219
    On Error GoTo vla_fail ' vla:613 src:219
    vla_step = 118 ' vla:614
    If vlatraceon() Then ' vla:615
        Call vlatracestep(118, vla_step_text(118)) ' vla:615
    End If
    ' ---- instructions.txt:222 Go to sheet Output. ----
    Call worksheets("output").activate
    vla_step = 119 ' vla:619
    If vlatraceon() Then ' vla:620
        Call vlatracestep(119, vla_step_text(119)) ' vla:620
    End If
    range("'Q1 Data'!A1") = "spaced" ' vla:622 src:223
    vla_step = 120 ' vla:623
    If vlatraceon() Then ' vla:624
        Call vlatracestep(120, vla_step_text(120)) ' vla:624
    End If
    Dim spaced_check As String ' vla:626 src:224
    vla_step = 121 ' vla:627
    If vlatraceon() Then ' vla:628
        Call vlatracestep(121, vla_step_text(121)) ' vla:628
    End If
    spaced_check = range("'Q1 Data'!A1") ' vla:630 src:225
    vla_step = 123 ' vla:631
    If vlatraceon() Then ' vla:632
        Call vlatracestep(123, vla_step_text(123)) ' vla:632
    End If
    ' ---- instructions.txt:232 Create a number called growth-check. ----
    Dim growth_check As Double ' vla:635 src:237
    vla_step = 124 ' vla:636
    If vlatraceon() Then ' vla:637
        Call vlatracestep(124, vla_step_text(124)) ' vla:637
    End If
    growth_check = 200 ' vla:639 src:238
    vla_step = 125 ' vla:640
    If vlatraceon() Then ' vla:641
        Call vlatracestep(125, vla_step_text(125)) ' vla:641
    End If
    growth_check = (growth_check * (1 + (10 / 100))) ' vla:643 src:239
    vla_step = 126 ' vla:644
    If vlatraceon() Then ' vla:645
        Call vlatracestep(126, vla_step_text(126)) ' vla:645
    End If
    growth_check = (growth_check * (1 + (50 / 100))) ' vla:647 src:240
    vla_step = 127 ' vla:648
    If vlatraceon() Then ' vla:649
        Call vlatracestep(127, vla_step_text(127)) ' vla:649
    End If
    growth_check = (growth_check * (1 + (100 / 100))) ' vla:651 src:241
    vla_step = 128 ' vla:652
    If vlatraceon() Then ' vla:653
        Call vlatracestep(128, vla_step_text(128)) ' vla:653
    End If
    growth_check = (growth_check * (1 - (75 / 100))) ' vla:655 src:242
    vla_step = 129 ' vla:656
    If vlatraceon() Then ' vla:657
        Call vlatracestep(129, vla_step_text(129)) ' vla:657
    End If
    pick_check = vlaitem(found_items, 2) ' vla:659 src:243
    vla_step = 130 ' vla:660
    If vlatraceon() Then ' vla:661
        Call vlatracestep(130, vla_step_text(130)) ' vla:661
    End If
    Debug.Print ((((("grew to " & growth_check) & ", picked ") & pick_check) & ", first ") & vlafirst(found_items)) ' vla:663 src:244
    vla_step = 131 ' vla:664
    If vlatraceon() Then ' vla:665
        Call vlatracestep(131, vla_step_text(131)) ' vla:665
    End If
    Dim quote_check As String ' vla:667 src:245
    vla_step = 132 ' vla:668
    If vlatraceon() Then ' vla:669
        Call vlatracestep(132, vla_step_text(132)) ' vla:669
    End If
    quote_check = "He said ""ok""" ' vla:671 src:246
    vla_step = 133 ' vla:672
    If vlatraceon() Then ' vla:673
        Call vlatracestep(133, vla_step_text(133)) ' vla:673
    End If
    ' ---- instructions.txt:248 Create a lookup called prices. ----
    Dim prices As Object ' vla:676 src:260
    Set prices = vladictnew() ' vla:676 src:260
    vla_step = 134 ' vla:677
    If vlatraceon() Then ' vla:678
        Call vlatracestep(134, vla_step_text(134)) ' vla:678
    End If
    Call vladictset(prices, "ax-7", 100) ' vla:680 src:261
    vla_step = 135 ' vla:681
    If vlatraceon() Then ' vla:682
        Call vlatracestep(135, vla_step_text(135)) ' vla:682
    End If
    Call vladictset(prices, "bx-2", 250) ' vla:684 src:262
    vla_step = 136 ' vla:685
    If vlatraceon() Then ' vla:686
        Call vlatracestep(136, vla_step_text(136)) ' vla:686
    End If
    Call vladictset(prices, "AX-7", 120) ' vla:688 src:263
    vla_step = 137 ' vla:689
    If vlatraceon() Then ' vla:690
        Call vlatracestep(137, vla_step_text(137)) ' vla:690
    End If
    ax_price = vladictget(prices, "ax-7") ' vla:692 src:264
    vla_step = 138 ' vla:693
    If vlatraceon() Then ' vla:694
        Call vlatracestep(138, vla_step_text(138)) ' vla:694
    End If
    key_count = vlacount(vladictkeys(prices)) ' vla:696 src:265
    vla_step = 139 ' vla:697
    If vlatraceon() Then ' vla:698
        Call vlatracestep(139, vla_step_text(139)) ' vla:698
    End If
    price_sum = (vladictget(prices, "ax-7") + vladictget(prices, "bx-2")) ' vla:700 src:266
    vla_step = 140 ' vla:701
    If vlatraceon() Then ' vla:702
        Call vlatracestep(140, vla_step_text(140)) ' vla:702
    End If
    Dim key_list As String ' vla:704 src:267
    vla_step = 141 ' vla:705
    If vlatraceon() Then ' vla:706
        Call vlatracestep(141, vla_step_text(141)) ' vla:706
    End If
    For Each k In vladictkeys(prices) ' vla:708 src:268
        vla_step = 142 ' vla:709 src:268
        If vlatraceon() Then ' vla:710 src:268
            Call vlatracestep(142, vla_step_text(142)) ' vla:710 src:268
        End If
        key_list = (key_list & k) ' vla:712 src:269
    Next k
    vla_step = 143 ' vla:713
    If vlatraceon() Then ' vla:714
        Call vlatracestep(143, vla_step_text(143)) ' vla:714
    End If
    ' ---- instructions.txt:271 Create a text called price-verdict. ----
    Dim price_verdict As String ' vla:717 src:271
    vla_step = 144 ' vla:718
    If vlatraceon() Then ' vla:719
        Call vlatracestep(144, vla_step_text(144)) ' vla:719
    End If
    If (vladictget(prices, "bx-2") > 200) Then ' vla:721 src:272
        price_verdict = "steep" ' vla:721 src:272
    End If
    vla_step = 145 ' vla:722
    If vlatraceon() Then ' vla:723
        Call vlatracestep(145, vla_step_text(145)) ' vla:723
    End If
    Dim pair_trace As String ' vla:725 src:273
    vla_step = 146 ' vla:726
    If vlatraceon() Then ' vla:727
        Call vlatracestep(146, vla_step_text(146)) ' vla:727
    End If
    For Each pair In vladictpairs(prices) ' vla:729 src:274
        vla_step = 147 ' vla:730 src:274
        If vlatraceon() Then ' vla:731 src:274
            Call vlatracestep(147, vla_step_text(147)) ' vla:731 src:274
        End If
        If (vlapairvalue(pair) > 200) Then ' vla:733 src:275
            pair_trace = (pair_trace & vlapairkey(pair)) ' vla:733 src:275
        End If
    Next pair
    vla_step = 148 ' vla:734
    If vlatraceon() Then ' vla:735
        Call vlatracestep(148, vla_step_text(148)) ' vla:735
    End If
    ' ---- instructions.txt:277 Log "lookup: ax " joined with ax-price joined with ", keys " joined... ----
    Debug.Print ((((("lookup: ax " & ax_price) & ", keys ") & key_list) & ", pairs ") & pair_trace) ' vla:738 src:277
    vla_step = 149 ' vla:739
    If vlatraceon() Then ' vla:740
        Call vlatracestep(149, vla_step_text(149)) ' vla:740
    End If
    ' ---- instructions.txt:279 Put "Item" into cell J1. ----
    range("j1") = "Item" ' vla:743 src:284
    vla_step = 150 ' vla:744
    If vlatraceon() Then ' vla:745
        Call vlatracestep(150, vla_step_text(150)) ' vla:745
    End If
    range("k1") = "Amount" ' vla:747 src:285
    vla_step = 151 ' vla:748
    If vlatraceon() Then ' vla:749
        Call vlatracestep(151, vla_step_text(151)) ' vla:749
    End If
    range("j2") = "Widget" ' vla:751 src:286
    vla_step = 152 ' vla:752
    If vlatraceon() Then ' vla:753
        Call vlatracestep(152, vla_step_text(152)) ' vla:753
    End If
    range("k2") = 10 ' vla:755 src:287
    vla_step = 153 ' vla:756
    If vlatraceon() Then ' vla:757
        Call vlatracestep(153, vla_step_text(153)) ' vla:757
    End If
    range("j3") = "Gadget" ' vla:759 src:288
    vla_step = 154 ' vla:760
    If vlatraceon() Then ' vla:761
        Call vlatracestep(154, vla_step_text(154)) ' vla:761
    End If
    range("k3") = 20 ' vla:763 src:289
    vla_step = 155 ' vla:764
    If vlatraceon() Then ' vla:765
        Call vlatracestep(155, vla_step_text(155)) ' vla:765
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j1:k3"), xllistobjecthasheaders:=xlyes).name = "salestable"
    vla_step = 156 ' vla:768
    If vlatraceon() Then ' vla:769
        Call vlatracestep(156, vla_step_text(156)) ' vla:769
    End If
    activesheet.listobjects("salestable").tablestyle = "tablestylemedium9"
    vla_step = 157 ' vla:772
    If vlatraceon() Then ' vla:773
        Call vlatracestep(157, vla_step_text(157)) ' vla:773
    End If
    activesheet.listobjects("salestable").showtotals = True
    vla_step = 158 ' vla:776
    If vlatraceon() Then ' vla:777
        Call vlatracestep(158, vla_step_text(158)) ' vla:777
    End If
    ' ---- instructions.txt:294 Put "X" into cell J5. ----
    range("j5") = "X" ' vla:780 src:296
    vla_step = 159 ' vla:781
    If vlatraceon() Then ' vla:782
        Call vlatracestep(159, vla_step_text(159)) ' vla:782
    End If
    range("j6") = "Y" ' vla:784 src:297
    vla_step = 160 ' vla:785
    If vlatraceon() Then ' vla:786
        Call vlatracestep(160, vla_step_text(160)) ' vla:786
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j5:j6"), xllistobjecthasheaders:=xlyes).name = "quiettable"
    vla_step = 161 ' vla:789
    If vlatraceon() Then ' vla:790
        Call vlatracestep(161, vla_step_text(161)) ' vla:790
    End If
    activesheet.listobjects("quiettable").showtotals = True
    vla_step = 162 ' vla:793
    If vlatraceon() Then ' vla:794
        Call vlatracestep(162, vla_step_text(162)) ' vla:794
    End If
    activesheet.listobjects("quiettable").showtotals = False
    vla_step = 163 ' vla:797
    If vlatraceon() Then ' vla:798
        Call vlatracestep(163, vla_step_text(163)) ' vla:798
    End If
    ' ---- instructions.txt:302 Put "A" into cell J8. ----
    range("j8") = "A" ' vla:801 src:304
    vla_step = 164 ' vla:802
    If vlatraceon() Then ' vla:803
        Call vlatracestep(164, vla_step_text(164)) ' vla:803
    End If
    range("j9") = "B" ' vla:805 src:305
    vla_step = 165 ' vla:806
    If vlatraceon() Then ' vla:807
        Call vlatracestep(165, vla_step_text(165)) ' vla:807
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j8:j9"), xllistobjecthasheaders:=xlyes).name = "temptable"
    vla_step = 166 ' vla:810
    If vlatraceon() Then ' vla:811
        Call vlatracestep(166, vla_step_text(166)) ' vla:811
    End If
    Call activesheet.listobjects("temptable").unlist
    vla_step = 167 ' vla:814
    If vlatraceon() Then ' vla:815
        Call vlatracestep(167, vla_step_text(167)) ' vla:815
    End If
    ' ---- instructions.txt:309 Put "Item" into cell J11. ----
    range("j11") = "Item" ' vla:818 src:312
    vla_step = 168 ' vla:819
    If vlatraceon() Then ' vla:820
        Call vlatracestep(168, vla_step_text(168)) ' vla:820
    End If
    range("k11") = "Qty" ' vla:822 src:313
    vla_step = 169 ' vla:823
    If vlatraceon() Then ' vla:824
        Call vlatracestep(169, vla_step_text(169)) ' vla:824
    End If
    range("j12") = "Bolt" ' vla:826 src:314
    vla_step = 170 ' vla:827
    If vlatraceon() Then ' vla:828
        Call vlatracestep(170, vla_step_text(170)) ' vla:828
    End If
    range("k12") = 5 ' vla:830 src:315
    vla_step = 171 ' vla:831
    If vlatraceon() Then ' vla:832
        Call vlatracestep(171, vla_step_text(171)) ' vla:832
    End If
    range("j13") = "Nut" ' vla:834 src:316
    vla_step = 172 ' vla:835
    If vlatraceon() Then ' vla:836
        Call vlatracestep(172, vla_step_text(172)) ' vla:836
    End If
    range("k13") = 8 ' vla:838 src:317
    vla_step = 173 ' vla:839
    If vlatraceon() Then ' vla:840
        Call vlatracestep(173, vla_step_text(173)) ' vla:840
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j11:k13"), xllistobjecthasheaders:=xlyes).name = "edittable"
    vla_step = 174 ' vla:843
    If vlatraceon() Then ' vla:844
        Call vlatracestep(174, vla_step_text(174)) ' vla:844
    End If
    Call activesheet.listobjects("edittable").listrows.add
    vla_step = 175 ' vla:847
    If vlatraceon() Then ' vla:848
        Call vlatracestep(175, vla_step_text(175)) ' vla:848
    End If
    Call activesheet.listobjects("edittable").listrows(1).delete
    vla_step = 176 ' vla:851
    If vlatraceon() Then ' vla:852
        Call vlatracestep(176, vla_step_text(176)) ' vla:852
    End If
    Set qty_values = activesheet.listobjects("edittable").listcolumns("qty").databodyrange
    vla_step = 177 ' vla:855
    If vlatraceon() Then ' vla:856
        Call vlatracestep(177, vla_step_text(177)) ' vla:856
    End If
    For Each q In qty_values ' vla:858 src:322
        Debug.Print q ' vla:858 src:322
    Next q
    vla_step = 178 ' vla:859
    If vlatraceon() Then ' vla:860
        Call vlatracestep(178, vla_step_text(178)) ' vla:860
    End If
    ' ---- instructions.txt:324 Put "Region" into cell N1. ----
    range("n1") = "Region" ' vla:863 src:332
    vla_step = 179 ' vla:864
    If vlatraceon() Then ' vla:865
        Call vlatracestep(179, vla_step_text(179)) ' vla:865
    End If
    range("o1") = "Product" ' vla:867 src:333
    vla_step = 180 ' vla:868
    If vlatraceon() Then ' vla:869
        Call vlatracestep(180, vla_step_text(180)) ' vla:869
    End If
    range("p1") = "Segment" ' vla:871 src:334
    vla_step = 181 ' vla:872
    If vlatraceon() Then ' vla:873
        Call vlatracestep(181, vla_step_text(181)) ' vla:873
    End If
    range("q1") = "Channel" ' vla:875 src:335
    vla_step = 182 ' vla:876
    If vlatraceon() Then ' vla:877
        Call vlatracestep(182, vla_step_text(182)) ' vla:877
    End If
    range("r1") = "Units" ' vla:879 src:336
    vla_step = 183 ' vla:880
    If vlatraceon() Then ' vla:881
        Call vlatracestep(183, vla_step_text(183)) ' vla:881
    End If
    range("s1") = "Revenue" ' vla:883 src:337
    vla_step = 184 ' vla:884
    If vlatraceon() Then ' vla:885
        Call vlatracestep(184, vla_step_text(184)) ' vla:885
    End If
    range("n2") = "North" ' vla:887 src:338
    vla_step = 185 ' vla:888
    If vlatraceon() Then ' vla:889
        Call vlatracestep(185, vla_step_text(185)) ' vla:889
    End If
    range("o2") = "Widget" ' vla:891 src:339
    vla_step = 186 ' vla:892
    If vlatraceon() Then ' vla:893
        Call vlatracestep(186, vla_step_text(186)) ' vla:893
    End If
    range("p2") = "Retail" ' vla:895 src:340
    vla_step = 187 ' vla:896
    If vlatraceon() Then ' vla:897
        Call vlatracestep(187, vla_step_text(187)) ' vla:897
    End If
    range("q2") = "Online" ' vla:899 src:341
    vla_step = 188 ' vla:900
    If vlatraceon() Then ' vla:901
        Call vlatracestep(188, vla_step_text(188)) ' vla:901
    End If
    range("r2") = 10 ' vla:903 src:342
    vla_step = 189 ' vla:904
    If vlatraceon() Then ' vla:905
        Call vlatracestep(189, vla_step_text(189)) ' vla:905
    End If
    range("s2") = 500 ' vla:907 src:343
    vla_step = 190 ' vla:908
    If vlatraceon() Then ' vla:909
        Call vlatracestep(190, vla_step_text(190)) ' vla:909
    End If
    range("n3") = "North" ' vla:911 src:344
    vla_step = 191 ' vla:912
    If vlatraceon() Then ' vla:913
        Call vlatracestep(191, vla_step_text(191)) ' vla:913
    End If
    range("o3") = "Gadget" ' vla:915 src:345
    vla_step = 192 ' vla:916
    If vlatraceon() Then ' vla:917
        Call vlatracestep(192, vla_step_text(192)) ' vla:917
    End If
    range("p3") = "Wholesale" ' vla:919 src:346
    vla_step = 193 ' vla:920
    If vlatraceon() Then ' vla:921
        Call vlatracestep(193, vla_step_text(193)) ' vla:921
    End If
    range("q3") = "Store" ' vla:923 src:347
    vla_step = 194 ' vla:924
    If vlatraceon() Then ' vla:925
        Call vlatracestep(194, vla_step_text(194)) ' vla:925
    End If
    range("r3") = 5 ' vla:927 src:348
    vla_step = 195 ' vla:928
    If vlatraceon() Then ' vla:929
        Call vlatracestep(195, vla_step_text(195)) ' vla:929
    End If
    range("s3") = 200 ' vla:931 src:349
    vla_step = 196 ' vla:932
    If vlatraceon() Then ' vla:933
        Call vlatracestep(196, vla_step_text(196)) ' vla:933
    End If
    range("n4") = "South" ' vla:935 src:350
    vla_step = 197 ' vla:936
    If vlatraceon() Then ' vla:937
        Call vlatracestep(197, vla_step_text(197)) ' vla:937
    End If
    range("o4") = "Widget" ' vla:939 src:351
    vla_step = 198 ' vla:940
    If vlatraceon() Then ' vla:941
        Call vlatracestep(198, vla_step_text(198)) ' vla:941
    End If
    range("p4") = "Wholesale" ' vla:943 src:352
    vla_step = 199 ' vla:944
    If vlatraceon() Then ' vla:945
        Call vlatracestep(199, vla_step_text(199)) ' vla:945
    End If
    range("q4") = "Online" ' vla:947 src:353
    vla_step = 200 ' vla:948
    If vlatraceon() Then ' vla:949
        Call vlatracestep(200, vla_step_text(200)) ' vla:949
    End If
    range("r4") = 20 ' vla:951 src:354
    vla_step = 201 ' vla:952
    If vlatraceon() Then ' vla:953
        Call vlatracestep(201, vla_step_text(201)) ' vla:953
    End If
    range("s4") = 900 ' vla:955 src:355
    vla_step = 202 ' vla:956
    If vlatraceon() Then ' vla:957
        Call vlatracestep(202, vla_step_text(202)) ' vla:957
    End If
    range("n5") = "South" ' vla:959 src:356
    vla_step = 203 ' vla:960
    If vlatraceon() Then ' vla:961
        Call vlatracestep(203, vla_step_text(203)) ' vla:961
    End If
    range("o5") = "Gadget" ' vla:963 src:357
    vla_step = 204 ' vla:964
    If vlatraceon() Then ' vla:965
        Call vlatracestep(204, vla_step_text(204)) ' vla:965
    End If
    range("p5") = "Retail" ' vla:967 src:358
    vla_step = 205 ' vla:968
    If vlatraceon() Then ' vla:969
        Call vlatracestep(205, vla_step_text(205)) ' vla:969
    End If
    range("q5") = "Store" ' vla:971 src:359
    vla_step = 206 ' vla:972
    If vlatraceon() Then ' vla:973
        Call vlatracestep(206, vla_step_text(206)) ' vla:973
    End If
    range("r5") = 8 ' vla:975 src:360
    vla_step = 207 ' vla:976
    If vlatraceon() Then ' vla:977
        Call vlatracestep(207, vla_step_text(207)) ' vla:977
    End If
    range("s5") = 300 ' vla:979 src:361
    vla_step = 208 ' vla:980
    If vlatraceon() Then ' vla:981
        Call vlatracestep(208, vla_step_text(208)) ' vla:981
    End If
    range("n6") = "East" ' vla:983 src:362
    vla_step = 209 ' vla:984
    If vlatraceon() Then ' vla:985
        Call vlatracestep(209, vla_step_text(209)) ' vla:985
    End If
    range("o6") = "Widget" ' vla:987 src:363
    vla_step = 210 ' vla:988
    If vlatraceon() Then ' vla:989
        Call vlatracestep(210, vla_step_text(210)) ' vla:989
    End If
    range("p6") = "Retail" ' vla:991 src:364
    vla_step = 211 ' vla:992
    If vlatraceon() Then ' vla:993
        Call vlatracestep(211, vla_step_text(211)) ' vla:993
    End If
    range("q6") = "Online" ' vla:995 src:365
    vla_step = 212 ' vla:996
    If vlatraceon() Then ' vla:997
        Call vlatracestep(212, vla_step_text(212)) ' vla:997
    End If
    range("r6") = 12 ' vla:999 src:366
    vla_step = 213 ' vla:1000
    If vlatraceon() Then ' vla:1001
        Call vlatracestep(213, vla_step_text(213)) ' vla:1001
    End If
    range("s6") = 600 ' vla:1003 src:367
    vla_step = 214 ' vla:1004
    If vlatraceon() Then ' vla:1005
        Call vlatracestep(214, vla_step_text(214)) ' vla:1005
    End If
    range("n7") = "East" ' vla:1007 src:368
    vla_step = 215 ' vla:1008
    If vlatraceon() Then ' vla:1009
        Call vlatracestep(215, vla_step_text(215)) ' vla:1009
    End If
    range("o7") = "Gadget" ' vla:1011 src:369
    vla_step = 216 ' vla:1012
    If vlatraceon() Then ' vla:1013
        Call vlatracestep(216, vla_step_text(216)) ' vla:1013
    End If
    range("p7") = "Wholesale" ' vla:1015 src:370
    vla_step = 217 ' vla:1016
    If vlatraceon() Then ' vla:1017
        Call vlatracestep(217, vla_step_text(217)) ' vla:1017
    End If
    range("q7") = "Store" ' vla:1019 src:371
    vla_step = 218 ' vla:1020
    If vlatraceon() Then ' vla:1021
        Call vlatracestep(218, vla_step_text(218)) ' vla:1021
    End If
    range("r7") = 6 ' vla:1023 src:372
    vla_step = 219 ' vla:1024
    If vlatraceon() Then ' vla:1025
        Call vlatracestep(219, vla_step_text(219)) ' vla:1025
    End If
    range("s7") = 250 ' vla:1027 src:373
    vla_step = 220 ' vla:1028
    If vlatraceon() Then ' vla:1029
        Call vlatracestep(220, vla_step_text(220)) ' vla:1029
    End If
    ' ---- instructions.txt:375 Make a pivot table from N1:S7 at U1 called SalesPivot. ----
    Call vlapivotcreate(range("n1:s7"), range("u1"), "salespivot")
    vla_step = 221 ' vla:1033
    If vlatraceon() Then ' vla:1034
        Call vlatracestep(221, vla_step_text(221)) ' vla:1034
    End If
    Call vlapivotsetorientation("salespivot", Array("region", "product"), "row")
    vla_step = 222 ' vla:1037
    If vlatraceon() Then ' vla:1038
        Call vlatracestep(222, vla_step_text(222)) ' vla:1038
    End If
    Call vlapivotsetorientation("salespivot", Array("segment"), "column")
    vla_step = 223 ' vla:1041
    If vlatraceon() Then ' vla:1042
        Call vlatracestep(223, vla_step_text(223)) ' vla:1042
    End If
    Call vlapivotsetorientation("salespivot", Array("channel"), "filter")
    vla_step = 224 ' vla:1045
    If vlatraceon() Then ' vla:1046
        Call vlatracestep(224, vla_step_text(224)) ' vla:1046
    End If
    ' ---- instructions.txt:386 Add Revenue to pivot SalesPivot as a sum. ----
    Call vlapivotaddvalues("salespivot", Array("revenue"), "sum")
    vla_step = 225 ' vla:1050
    If vlatraceon() Then ' vla:1051
        Call vlatracestep(225, vla_step_text(225)) ' vla:1051
    End If
    Call vlapivotaddvalues("salespivot", Array("revenue", "units"), "count")
    vla_step = 226 ' vla:1054
    If vlatraceon() Then ' vla:1055
        Call vlatracestep(226, vla_step_text(226)) ' vla:1055
    End If
    Call vlapivotaddvalues("salespivot", Array("units"), "average")
    vla_step = 227 ' vla:1058
    If vlatraceon() Then ' vla:1059
        Call vlatracestep(227, vla_step_text(227)) ' vla:1059
    End If
    ' ---- instructions.txt:396 Make a pivot table from N1:S7 at N20 called FullPivot with rows of ... ----
    Call vlapivotcreate(range("n1:s7"), range("n20"), "fullpivot")
    Call vlapivotsetorientation("fullpivot", Array("region", "product", "channel"), "row")
    Call vlapivotsetorientation("fullpivot", Array("segment"), "column")
    Call vlapivotaddvalues("fullpivot", Array("revenue", "units"), "sum")
    vla_step = 228 ' vla:1063
    If vlatraceon() Then ' vla:1064
        Call vlatracestep(228, vla_step_text(228)) ' vla:1064
    End If
    ' ---- instructions.txt:408 Refresh pivot SalesPivot. ----
    Call vlapivotrefresh("salespivot")
    vla_step = 229 ' vla:1068
    If vlatraceon() Then ' vla:1069
        Call vlatracestep(229, vla_step_text(229)) ' vla:1069
    End If
    Call vlapivotrefreshall
    vla_step = 230 ' vla:1072
    If vlatraceon() Then ' vla:1073
        Call vlatracestep(230, vla_step_text(230)) ' vla:1073
    End If
    Call vlapivotsetshowdetail("salespivot", Array("region"), False)
    vla_step = 231 ' vla:1076
    If vlatraceon() Then ' vla:1077
        Call vlatracestep(231, vla_step_text(231)) ' vla:1077
    End If
    ' ---- instructions.txt:419 Collapse Product in pivot FullPivot. ----
    Call vlapivotsetshowdetail("fullpivot", Array("product"), False)
    vla_step = 232 ' vla:1081
    If vlatraceon() Then ' vla:1082
        Call vlatracestep(232, vla_step_text(232)) ' vla:1082
    End If
    Call vlapivotsetshowdetail("fullpivot", Array("product"), True)
    vla_step = 233 ' vla:1085
    If vlatraceon() Then ' vla:1086
        Call vlatracestep(233, vla_step_text(233)) ' vla:1086
    End If
    ' ---- instructions.txt:430 Show pivot SalesPivot in tabular form. ----
    Call vlapivotsetrowlayout("salespivot", "tabular")
    vla_step = 234 ' vla:1090
    If vlatraceon() Then ' vla:1091
        Call vlatracestep(234, vla_step_text(234)) ' vla:1091
    End If
    Call vlapivotsetrowlayout("salespivot", "compact")
    vla_step = 235 ' vla:1094
    If vlatraceon() Then ' vla:1095
        Call vlatracestep(235, vla_step_text(235)) ' vla:1095
    End If
    Call vlapivotsetrowlayout("fullpivot", "outline")
    vla_step = 236 ' vla:1098
    If vlatraceon() Then ' vla:1099
        Call vlatracestep(236, vla_step_text(236)) ' vla:1099
    End If
    ' ---- instructions.txt:447 Hide subtotals for Region in pivot SalesPivot. ----
    Call vlapivotsetsubtotals("salespivot", Array("region"), False)
    vla_step = 237 ' vla:1103
    If vlatraceon() Then ' vla:1104
        Call vlatracestep(237, vla_step_text(237)) ' vla:1104
    End If
    Call vlapivotsetsubtotals("fullpivot", Array("product"), False)
    vla_step = 238 ' vla:1107
    If vlatraceon() Then ' vla:1108
        Call vlatracestep(238, vla_step_text(238)) ' vla:1108
    End If
    Call vlapivotsetsubtotals("fullpivot", Array("product"), True)
    vla_step = 239 ' vla:1111
    If vlatraceon() Then ' vla:1112
        Call vlatracestep(239, vla_step_text(239)) ' vla:1112
    End If
    ' ---- instructions.txt:457 Add a blank row after Region in pivot SalesPivot. ----
    Call vlapivotsetblankline("salespivot", Array("region"), True)
    vla_step = 240 ' vla:1116
    If vlatraceon() Then ' vla:1117
        Call vlatracestep(240, vla_step_text(240)) ' vla:1117
    End If
    Call vlapivotsetblankline("fullpivot", Array("product"), True)
    vla_step = 241 ' vla:1120
    If vlatraceon() Then ' vla:1121
        Call vlatracestep(241, vla_step_text(241)) ' vla:1121
    End If
    Call vlapivotsetblankline("fullpivot", Array("product"), False)
    vla_step = 242 ' vla:1124
    If vlatraceon() Then ' vla:1125
        Call vlatracestep(242, vla_step_text(242)) ' vla:1125
    End If
    ' ---- instructions.txt:461 Sort Region in pivot SalesPivot descending. ----
    Call vlapivotsort("salespivot", "region", "descending", "")
    vla_step = 243 ' vla:1129
    If vlatraceon() Then ' vla:1130
        Call vlatracestep(243, vla_step_text(243)) ' vla:1130
    End If
    Call vlapivotsort("salespivot", "region", "ascending", "")
    vla_step = 244 ' vla:1133
    If vlatraceon() Then ' vla:1134
        Call vlatracestep(244, vla_step_text(244)) ' vla:1134
    End If
    Call vlapivotsort("fullpivot", "product", "descending", "revenue")
    vla_step = 245 ' vla:1137
    If vlatraceon() Then ' vla:1138
        Call vlatracestep(245, vla_step_text(245)) ' vla:1138
    End If
    Call vlapivotsort("fullpivot", "product", "ascending", "revenue")
    vla_step = 246 ' vla:1141
    If vlatraceon() Then ' vla:1142
        Call vlatracestep(246, vla_step_text(246)) ' vla:1142
    End If
    ' ---- instructions.txt:481 Make a pivot table from N1:S7 at N200 called RenameMePivot. ----
    Call vlapivotcreate(range("n1:s7"), range("n200"), "renamemepivot")
    vla_step = 247 ' vla:1146
    If vlatraceon() Then ' vla:1147
        Call vlatracestep(247, vla_step_text(247)) ' vla:1147
    End If
    Call vlapivotrename("renamemepivot", "renamedpivot")
    vla_step = 248 ' vla:1150
    If vlatraceon() Then ' vla:1151
        Call vlatracestep(248, vla_step_text(248)) ' vla:1151
    End If
    ' ---- instructions.txt:493 Make a pivot table from N1:S7 at N220 called ClearMePivot. ----
    Call vlapivotcreate(range("n1:s7"), range("n220"), "clearmepivot")
    vla_step = 249 ' vla:1155
    If vlatraceon() Then ' vla:1156
        Call vlatracestep(249, vla_step_text(249)) ' vla:1156
    End If
    Call vlapivotsetorientation("clearmepivot", Array("region"), "row")
    vla_step = 250 ' vla:1159
    If vlatraceon() Then ' vla:1160
        Call vlatracestep(250, vla_step_text(250)) ' vla:1160
    End If
    Call vlapivotaddvalues("clearmepivot", Array("revenue"), "sum")
    vla_step = 251 ' vla:1163
    If vlatraceon() Then ' vla:1164
        Call vlatracestep(251, vla_step_text(251)) ' vla:1164
    End If
    Call vlapivotclear("clearmepivot")
    vla_step = 252 ' vla:1167
    If vlatraceon() Then ' vla:1168
        Call vlatracestep(252, vla_step_text(252)) ' vla:1168
    End If
    ' ---- instructions.txt:498 Make a pivot table from N1:S7 at N240 called RemoveFieldMePivot. ----
    Call vlapivotcreate(range("n1:s7"), range("n240"), "removefieldmepivot")
    vla_step = 253 ' vla:1172
    If vlatraceon() Then ' vla:1173
        Call vlatracestep(253, vla_step_text(253)) ' vla:1173
    End If
    Call vlapivotsetorientation("removefieldmepivot", Array("region", "product"), "row")
    vla_step = 254 ' vla:1176
    If vlatraceon() Then ' vla:1177
        Call vlatracestep(254, vla_step_text(254)) ' vla:1177
    End If
    Call vlapivotsetorientation("removefieldmepivot", Array("product"), "hidden")
    vla_step = 255 ' vla:1180
    If vlatraceon() Then ' vla:1181
        Call vlatracestep(255, vla_step_text(255)) ' vla:1181
    End If
    ' ---- instructions.txt:502 Put "West" into cell N8. ----
    range("n8") = "West" ' vla:1184 src:506
    vla_step = 256 ' vla:1185
    If vlatraceon() Then ' vla:1186
        Call vlatracestep(256, vla_step_text(256)) ' vla:1186
    End If
    range("o8") = "Widget" ' vla:1188 src:507
    vla_step = 257 ' vla:1189
    If vlatraceon() Then ' vla:1190
        Call vlatracestep(257, vla_step_text(257)) ' vla:1190
    End If
    range("p8") = "Retail" ' vla:1192 src:508
    vla_step = 258 ' vla:1193
    If vlatraceon() Then ' vla:1194
        Call vlatracestep(258, vla_step_text(258)) ' vla:1194
    End If
    range("q8") = "Online" ' vla:1196 src:509
    vla_step = 259 ' vla:1197
    If vlatraceon() Then ' vla:1198
        Call vlatracestep(259, vla_step_text(259)) ' vla:1198
    End If
    range("r8") = 15 ' vla:1200 src:510
    vla_step = 260 ' vla:1201
    If vlatraceon() Then ' vla:1202
        Call vlatracestep(260, vla_step_text(260)) ' vla:1202
    End If
    range("s8") = 700 ' vla:1204 src:511
    vla_step = 261 ' vla:1205
    If vlatraceon() Then ' vla:1206
        Call vlatracestep(261, vla_step_text(261)) ' vla:1206
    End If
    ' ---- instructions.txt:513 Make a pivot table from N1:S7 at N260 called SourceTestPivot. ----
    Call vlapivotcreate(range("n1:s7"), range("n260"), "sourcetestpivot")
    vla_step = 262 ' vla:1210
    If vlatraceon() Then ' vla:1211
        Call vlatracestep(262, vla_step_text(262)) ' vla:1211
    End If
    Call vlapivotsetorientation("sourcetestpivot", Array("region"), "row")
    vla_step = 263 ' vla:1214
    If vlatraceon() Then ' vla:1215
        Call vlatracestep(263, vla_step_text(263)) ' vla:1215
    End If
    Call vlapivotchangesource("sourcetestpivot", range("n1:s8"))
    vla_step = 264 ' vla:1218
    If vlatraceon() Then ' vla:1219
        Call vlatracestep(264, vla_step_text(264)) ' vla:1219
    End If
    ' ---- instructions.txt:517 Make a pivot table from N1:S7 at N40 called TempPivot. ----
    Call vlapivotcreate(range("n1:s7"), range("n40"), "temppivot")
    vla_step = 265 ' vla:1223
    If vlatraceon() Then ' vla:1224
        Call vlatracestep(265, vla_step_text(265)) ' vla:1224
    End If
    Call vlapivotdelete("temppivot")
    vla_step = 266 ' vla:1227
    If vlatraceon() Then ' vla:1228
        Call vlatracestep(266, vla_step_text(266)) ' vla:1228
    End If
    ' ---- instructions.txt:523 Put grand into cell H1. ----
    range("h1") = grand ' vla:1231 src:524
    vla_step = 267 ' vla:1232
    If vlatraceon() Then ' vla:1233
        Call vlatracestep(267, vla_step_text(267)) ' vla:1233
    End If
    range("h2") = biggest ' vla:1235 src:525
    vla_step = 268 ' vla:1236
    If vlatraceon() Then ' vla:1237
        Call vlatracestep(268, vla_step_text(268)) ' vla:1237
    End If
    range("h3") = round_check ' vla:1239 src:526
    vla_step = 269 ' vla:1240
    If vlatraceon() Then ' vla:1241
        Call vlatracestep(269, vla_step_text(269)) ' vla:1241
    End If
    range("h4") = thousand_check ' vla:1243 src:527
    vla_step = 270 ' vla:1244
    If vlatraceon() Then ' vla:1245
        Call vlatracestep(270, vla_step_text(270)) ' vla:1245
    End If
    range("h5") = search_row ' vla:1247 src:528
    vla_step = 271 ' vla:1248
    If vlatraceon() Then ' vla:1249
        Call vlatracestep(271, vla_step_text(271)) ' vla:1249
    End If
    range("h6") = f_last ' vla:1251 src:529
    vla_step = 272 ' vla:1252
    If vlatraceon() Then ' vla:1253
        Call vlatracestep(272, vla_step_text(272)) ' vla:1253
    End If
    range("h7") = list_count ' vla:1255 src:530
    vla_step = 273 ' vla:1256
    If vlatraceon() Then ' vla:1257
        Call vlatracestep(273, vla_step_text(273)) ' vla:1257
    End If
    range("h8") = verdict ' vla:1259 src:531
    vla_step = 274 ' vla:1260
    If vlatraceon() Then ' vla:1261
        Call vlatracestep(274, vla_step_text(274)) ' vla:1261
    End If
    range("h9") = region_label ' vla:1263 src:532
    vla_step = 275 ' vla:1264
    If vlatraceon() Then ' vla:1265
        Call vlatracestep(275, vla_step_text(275)) ' vla:1265
    End If
    range("h10") = until_count ' vla:1267 src:533
    vla_step = 276 ' vla:1268
    If vlatraceon() Then ' vla:1269
        Call vlatracestep(276, vla_step_text(276)) ' vla:1269
    End If
    range("h11") = rescue ' vla:1271 src:534
    vla_step = 277 ' vla:1272
    If vlatraceon() Then ' vla:1273
        Call vlatracestep(277, vla_step_text(277)) ' vla:1273
    End If
    range("h12") = risk_free ' vla:1275 src:535
    vla_step = 278 ' vla:1276
    If vlatraceon() Then ' vla:1277
        Call vlatracestep(278, vla_step_text(278)) ' vla:1277
    End If
    range("h13") = fee ' vla:1279 src:536
    vla_step = 279 ' vla:1280
    If vlatraceon() Then ' vla:1281
        Call vlatracestep(279, vla_step_text(279)) ' vla:1281
    End If
    range("h14") = fee_size ' vla:1283 src:537
    vla_step = 280 ' vla:1284
    If vlatraceon() Then ' vla:1285
        Call vlatracestep(280, vla_step_text(280)) ' vla:1285
    End If
    range("h15") = num_col_check ' vla:1287 src:538
    vla_step = 281 ' vla:1288
    If vlatraceon() Then ' vla:1289
        Call vlatracestep(281, vla_step_text(281)) ' vla:1289
    End If
    range("h17") = value_word_check ' vla:1291 src:539
    vla_step = 282 ' vla:1292
    If vlatraceon() Then ' vla:1293
        Call vlatracestep(282, vla_step_text(282)) ' vla:1293
    End If
    range("h18") = spaced_check ' vla:1295 src:540
    vla_step = 283 ' vla:1296
    If vlatraceon() Then ' vla:1297
        Call vlatracestep(283, vla_step_text(283)) ' vla:1297
    End If
    range("h19") = full_commission ' vla:1299 src:541
    vla_step = 284 ' vla:1300
    If vlatraceon() Then ' vla:1301
        Call vlatracestep(284, vla_step_text(284)) ' vla:1301
    End If
    range("h20") = default_commission ' vla:1303 src:542
    vla_step = 285 ' vla:1304
    If vlatraceon() Then ' vla:1305
        Call vlatracestep(285, vla_step_text(285)) ' vla:1305
    End If
    range("h21") = growth_check ' vla:1307 src:543
    vla_step = 286 ' vla:1308
    If vlatraceon() Then ' vla:1309
        Call vlatracestep(286, vla_step_text(286)) ' vla:1309
    End If
    range("h22") = pick_check ' vla:1311 src:544
    vla_step = 287 ' vla:1312
    If vlatraceon() Then ' vla:1313
        Call vlatracestep(287, vla_step_text(287)) ' vla:1313
    End If
    range("h23") = quote_check ' vla:1315 src:545
    vla_step = 288 ' vla:1316
    If vlatraceon() Then ' vla:1317
        Call vlatracestep(288, vla_step_text(288)) ' vla:1317
    End If
    range("h24") = vat_rate() ' vla:1319 src:546
    vla_step = 289 ' vla:1320
    If vlatraceon() Then ' vla:1321
        Call vlatracestep(289, vla_step_text(289)) ' vla:1321
    End If
    range("h26") = ax_price ' vla:1323 src:547
    vla_step = 290 ' vla:1324
    If vlatraceon() Then ' vla:1325
        Call vlatracestep(290, vla_step_text(290)) ' vla:1325
    End If
    range("h27") = key_count ' vla:1327 src:548
    vla_step = 291 ' vla:1328
    If vlatraceon() Then ' vla:1329
        Call vlatracestep(291, vla_step_text(291)) ' vla:1329
    End If
    range("h28") = key_list ' vla:1331 src:549
    vla_step = 292 ' vla:1332
    If vlatraceon() Then ' vla:1333
        Call vlatracestep(292, vla_step_text(292)) ' vla:1333
    End If
    range("h29") = price_sum ' vla:1335 src:550
    vla_step = 293 ' vla:1336
    If vlatraceon() Then ' vla:1337
        Call vlatracestep(293, vla_step_text(293)) ' vla:1337
    End If
    range("h30") = price_verdict ' vla:1339 src:551
    vla_step = 294 ' vla:1340
    If vlatraceon() Then ' vla:1341
        Call vlatracestep(294, vla_step_text(294)) ' vla:1341
    End If
    range("h31") = pair_trace ' vla:1343 src:552
    vla_step = 295 ' vla:1344
    If vlatraceon() Then ' vla:1345
        Call vlatracestep(295, vla_step_text(295)) ' vla:1345
    End If
    ' ---- instructions.txt:554 (set! (range "h25") "vla-row") ----
    range("h25") = "vla-row" ' vla:1348 src:559
    vla_step = 296 ' vla:1349
    If vlatraceon() Then ' vla:1350
        Call vlatracestep(296, vla_step_text(296)) ' vla:1350
    End If
    For counter = 1 To 2
        vla_step = 297 ' vla:1353 src:560
        If vlatraceon() Then ' vla:1354 src:560
            Call vlatracestep(297, vla_step_text(297)) ' vla:1354 src:560
        End If
        Debug.Print counter ' vla:1356 src:561
    Next counter
    vla_step = 298 ' vla:1357
    If vlatraceon() Then ' vla:1358
        Call vlatracestep(298, vla_step_text(298)) ' vla:1358
    End If
    ' ---- instructions.txt:563 Work on sheet Demo. ----
    Call vlaensuresheet("demo") ' vla:1361 src:576
    Call worksheets("demo").activate
    vla_step = 299 ' vla:1362
    If vlatraceon() Then ' vla:1363
        Call vlatracestep(299, vla_step_text(299)) ' vla:1363
    End If
    ' ---- instructions.txt:578 Make cell A1 italic. ----
    range("a1").font.italic = True
    vla_step = 300 ' vla:1367
    If vlatraceon() Then ' vla:1368
        Call vlatracestep(300, vla_step_text(300)) ' vla:1368
    End If
    range("a2").interior.color = vbred
    vla_step = 301 ' vla:1371
    If vlatraceon() Then ' vla:1372
        Call vlatracestep(301, vla_step_text(301)) ' vla:1372
    End If
    range("a3").interior.color = vbyellow
    vla_step = 302 ' vla:1375
    If vlatraceon() Then ' vla:1376
        Call vlatracestep(302, vla_step_text(302)) ' vla:1376
    End If
    range("a4").font.color = vlacolor(hot_pink)
    vla_step = 303 ' vla:1379
    If vlatraceon() Then ' vla:1380
        Call vlatracestep(303, vla_step_text(303)) ' vla:1380
    End If
    range("a5").interior.color = vlacolor("#FF69B4")
    vla_step = 304 ' vla:1383
    If vlatraceon() Then ' vla:1384
        Call vlatracestep(304, vla_step_text(304)) ' vla:1384
    End If
    range("a5").interior.colorindex = xlnone
    vla_step = 305 ' vla:1387
    If vlatraceon() Then ' vla:1388
        Call vlatracestep(305, vla_step_text(305)) ' vla:1388
    End If
    range("a1:e10").borders.linestyle = xlcontinuous
    vla_step = 306 ' vla:1391
    If vlatraceon() Then ' vla:1392
        Call vlatracestep(306, vla_step_text(306)) ' vla:1392
    End If
    range("b1").numberformat = "$#,##0.00"
    vla_step = 307 ' vla:1395
    If vlatraceon() Then ' vla:1396
        Call vlatracestep(307, vla_step_text(307)) ' vla:1396
    End If
    range("b2").numberformat = "0.0%"
    vla_step = 308 ' vla:1399
    If vlatraceon() Then ' vla:1400
        Call vlatracestep(308, vla_step_text(308)) ' vla:1400
    End If
    range("b3").numberformat = "mm/dd/yyyy"
    vla_step = 309 ' vla:1403
    If vlatraceon() Then ' vla:1404
        Call vlatracestep(309, vla_step_text(309)) ' vla:1404
    End If
    range("a1").horizontalalignment = xlcenter
    vla_step = 310 ' vla:1407
    If vlatraceon() Then ' vla:1408
        Call vlatracestep(310, vla_step_text(310)) ' vla:1408
    End If
    range("a1:e1").horizontalalignment = xlcenter
    vla_step = 311 ' vla:1411
    If vlatraceon() Then ' vla:1412
        Call vlatracestep(311, vla_step_text(311)) ' vla:1412
    End If
    range("a2:a5").horizontalalignment = xlleft
    vla_step = 312 ' vla:1415
    If vlatraceon() Then ' vla:1416
        Call vlatracestep(312, vla_step_text(312)) ' vla:1416
    End If
    range("a6").horizontalalignment = xlcenter
    vla_step = 313 ' vla:1419
    If vlatraceon() Then ' vla:1420
        Call vlatracestep(313, vla_step_text(313)) ' vla:1420
    End If
    range("c1:c5").wraptext = True
    vla_step = 314 ' vla:1423
    If vlatraceon() Then ' vla:1424
        Call vlatracestep(314, vla_step_text(314)) ' vla:1424
    End If
    range("c1:c5").wraptext = False
    vla_step = 315 ' vla:1427
    If vlatraceon() Then ' vla:1428
        Call vlatracestep(315, vla_step_text(315)) ' vla:1428
    End If
    Call range("d1:d3").merge
    vla_step = 316 ' vla:1431
    If vlatraceon() Then ' vla:1432
        Call vlatracestep(316, vla_step_text(316)) ' vla:1432
    End If
    Call range("d1:d3").unmerge
    vla_step = 317 ' vla:1435
    If vlatraceon() Then ' vla:1436
        Call vlatracestep(317, vla_step_text(317)) ' vla:1436
    End If
    ' ---- instructions.txt:598 Set height of row 1 to 30. ----
    rows(1).rowheight = 30
    vla_step = 318 ' vla:1440
    If vlatraceon() Then ' vla:1441
        Call vlatracestep(318, vla_step_text(318)) ' vla:1441
    End If
    columns("a").columnwidth = 20
    vla_step = 319 ' vla:1444
    If vlatraceon() Then ' vla:1445
        Call vlatracestep(319, vla_step_text(319)) ' vla:1445
    End If
    Call columns("b").insert
    vla_step = 320 ' vla:1448
    If vlatraceon() Then ' vla:1449
        Call vlatracestep(320, vla_step_text(320)) ' vla:1449
    End If
    columns("c").hidden = True
    vla_step = 321 ' vla:1452
    If vlatraceon() Then ' vla:1453
        Call vlatracestep(321, vla_step_text(321)) ' vla:1453
    End If
    columns("c").hidden = False
    vla_step = 322 ' vla:1456
    If vlatraceon() Then ' vla:1457
        Call vlatracestep(322, vla_step_text(322)) ' vla:1457
    End If
    Call columns("b").delete
    vla_step = 323 ' vla:1460
    If vlatraceon() Then ' vla:1461
        Call vlatracestep(323, vla_step_text(323)) ' vla:1461
    End If
    Call rows(5).insert
    vla_step = 324 ' vla:1464
    If vlatraceon() Then ' vla:1465
        Call vlatracestep(324, vla_step_text(324)) ' vla:1465
    End If
    Call rows(1).insert
    vla_step = 325 ' vla:1468
    If vlatraceon() Then ' vla:1469
        Call vlatracestep(325, vla_step_text(325)) ' vla:1469
    End If
    Call rows(2).delete
    vla_step = 326 ' vla:1472
    If vlatraceon() Then ' vla:1473
        Call vlatracestep(326, vla_step_text(326)) ' vla:1473
    End If
    Call rows(3).delete
    vla_step = 327 ' vla:1476
    If vlatraceon() Then ' vla:1477
        Call vlatracestep(327, vla_step_text(327)) ' vla:1477
    End If
    rows(10).hidden = True
    vla_step = 328 ' vla:1480
    If vlatraceon() Then ' vla:1481
        Call vlatracestep(328, vla_step_text(328)) ' vla:1481
    End If
    rows(10).hidden = False
    vla_step = 329 ' vla:1484
    If vlatraceon() Then ' vla:1485
        Call vlatracestep(329, vla_step_text(329)) ' vla:1485
    End If
    Call rows(2).select
    activewindow.freezepanes = True
    vla_step = 330 ' vla:1488
    If vlatraceon() Then ' vla:1489
        Call vlatracestep(330, vla_step_text(330)) ' vla:1489
    End If
    activewindow.freezepanes = False
    vla_step = 331 ' vla:1492
    If vlatraceon() Then ' vla:1493
        Call vlatracestep(331, vla_step_text(331)) ' vla:1493
    End If
    Call columns("a").autofit
    vla_step = 332 ' vla:1496
    If vlatraceon() Then ' vla:1497
        Call vlatracestep(332, vla_step_text(332)) ' vla:1497
    End If
    Call cells.entirecolumn.autofit
    vla_step = 333 ' vla:1500
    If vlatraceon() Then ' vla:1501
        Call vlatracestep(333, vla_step_text(333)) ' vla:1501
    End If
    ' ---- instructions.txt:618 Put "Region" into cell G1. ----
    range("g1") = "Region" ' vla:1504 src:620
    vla_step = 334 ' vla:1505
    If vlatraceon() Then ' vla:1506
        Call vlatracestep(334, vla_step_text(334)) ' vla:1506
    End If
    range("h1") = "Amount" ' vla:1508 src:621
    vla_step = 335 ' vla:1509
    If vlatraceon() Then ' vla:1510
        Call vlatracestep(335, vla_step_text(335)) ' vla:1510
    End If
    range("i1") = "Notes" ' vla:1512 src:622
    vla_step = 336 ' vla:1513
    If vlatraceon() Then ' vla:1514
        Call vlatracestep(336, vla_step_text(336)) ' vla:1514
    End If
    range("g2") = "West" ' vla:1516 src:623
    vla_step = 337 ' vla:1517
    If vlatraceon() Then ' vla:1518
        Call vlatracestep(337, vla_step_text(337)) ' vla:1518
    End If
    range("h2") = 100 ' vla:1520 src:624
    vla_step = 338 ' vla:1521
    If vlatraceon() Then ' vla:1522
        Call vlatracestep(338, vla_step_text(338)) ' vla:1522
    End If
    range("i2") = "ok" ' vla:1524 src:625
    vla_step = 339 ' vla:1525
    If vlatraceon() Then ' vla:1526
        Call vlatracestep(339, vla_step_text(339)) ' vla:1526
    End If
    range("g3") = "East" ' vla:1528 src:626
    vla_step = 340 ' vla:1529
    If vlatraceon() Then ' vla:1530
        Call vlatracestep(340, vla_step_text(340)) ' vla:1530
    End If
    range("h3") = 250 ' vla:1532 src:627
    vla_step = 341 ' vla:1533
    If vlatraceon() Then ' vla:1534
        Call vlatracestep(341, vla_step_text(341)) ' vla:1534
    End If
    range("i3") = "ok" ' vla:1536 src:628
    vla_step = 342 ' vla:1537
    If vlatraceon() Then ' vla:1538
        Call vlatracestep(342, vla_step_text(342)) ' vla:1538
    End If
    range("g4") = "West" ' vla:1540 src:629
    vla_step = 343 ' vla:1541
    If vlatraceon() Then ' vla:1542
        Call vlatracestep(343, vla_step_text(343)) ' vla:1542
    End If
    range("h4") = 100 ' vla:1544 src:630
    vla_step = 344 ' vla:1545
    If vlatraceon() Then ' vla:1546
        Call vlatracestep(344, vla_step_text(344)) ' vla:1546
    End If
    range("i4") = "dup" ' vla:1548 src:631
    vla_step = 345 ' vla:1549
    If vlatraceon() Then ' vla:1550
        Call vlatracestep(345, vla_step_text(345)) ' vla:1550
    End If
    Call range("g1:i4").sort(key1:=range("h1"), order1:=xldescending, header:=xlyes)
    vla_step = 346 ' vla:1553
    If vlatraceon() Then ' vla:1554
        Call vlatracestep(346, vla_step_text(346)) ' vla:1554
    End If
    Call range("g1:i4").autofilter(field:=1, criteria1:="West")
    vla_step = 347 ' vla:1557
    If vlatraceon() Then ' vla:1558
        Call vlatracestep(347, vla_step_text(347)) ' vla:1558
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 348 ' vla:1561
    If vlatraceon() Then ' vla:1562
        Call vlatracestep(348, vla_step_text(348)) ' vla:1562
    End If
    Call range("g1:i4").replace(what:="dup", replacement:="ok")
    vla_step = 349 ' vla:1565
    If vlatraceon() Then ' vla:1566
        Call vlatracestep(349, vla_step_text(349)) ' vla:1566
    End If
    Call range("g1:i4").removeduplicates(columns:=1, header:=xlyes)
    vla_step = 350 ' vla:1569
    If vlatraceon() Then ' vla:1570
        Call vlatracestep(350, vla_step_text(350)) ' vla:1570
    End If
    Call columns("i").replace(what:="ok", replacement:="fine")
    vla_step = 351 ' vla:1573
    If vlatraceon() Then ' vla:1574
        Call vlatracestep(351, vla_step_text(351)) ' vla:1574
    End If
    Call range("g1:i4").copy
    Call range("g1:i4").pastespecial(paste:=xlpastevalues)
    application.cutcopymode = False
    vla_step = 352 ' vla:1577
    If vlatraceon() Then ' vla:1578
        Call vlatracestep(352, vla_step_text(352)) ' vla:1578
    End If
    Call range("g1:i4").clearformats
    vla_step = 353 ' vla:1581
    If vlatraceon() Then ' vla:1582
        Call vlatracestep(353, vla_step_text(353)) ' vla:1582
    End If
    Call vlacheckrangename("demo_table")
    range("g1:i4").name = "demo_table"
    vla_step = 354 ' vla:1585
    If vlatraceon() Then ' vla:1586
        Call vlatracestep(354, vla_step_text(354)) ' vla:1586
    End If
    Call range("g1:i4").copy(destination:=range("k1:m4"))
    vla_step = 355 ' vla:1589
    If vlatraceon() Then ' vla:1590
        Call vlatracestep(355, vla_step_text(355)) ' vla:1590
    End If
    ' ---- instructions.txt:643 Set tab-color of sheet Demo to hot-pink. ----
    worksheets("demo").tab.color = vlacolor(hot_pink)
    vla_step = 356 ' vla:1594
    If vlatraceon() Then ' vla:1595
        Call vlatracestep(356, vla_step_text(356)) ' vla:1595
    End If
    Call activesheet.protect(password:="demo123")
    vla_step = 357 ' vla:1598
    If vlatraceon() Then ' vla:1599
        Call vlatracestep(357, vla_step_text(357)) ' vla:1599
    End If
    Call activesheet.unprotect(password:="demo123")
    vla_step = 358 ' vla:1602
    If vlatraceon() Then ' vla:1603
        Call vlatracestep(358, vla_step_text(358)) ' vla:1603
    End If
    ' ---- instructions.txt:649 Try: ----
    On Error GoTo vla_tryf_4 ' vla:1606 src:675
    vla_step = 359 ' vla:1607 src:675
    If vlatraceon() Then ' vla:1608 src:675
        Call vlatracestep(359, vla_step_text(359)) ' vla:1608 src:675
    End If
    application.displayalerts = False
    Call worksheets("gstruct").delete
    application.displayalerts = True
    GoTo vla_tryd_4 ' vla:1611 src:675
vla_tryf_4: ' vla:1612 src:675
    vla_problem = err.description ' vla:1613 src:675
    Resume vla_tryr_4 ' vla:1614 src:675
vla_tryr_4: ' vla:1615 src:675
    On Error GoTo vla_fail ' vla:1616 src:675
vla_tryd_4: ' vla:1617 src:675
    On Error GoTo vla_fail ' vla:1618 src:675
    vla_step = 360 ' vla:1619
    If vlatraceon() Then ' vla:1620
        Call vlatracestep(360, vla_step_text(360)) ' vla:1620
    End If
    ' ---- instructions.txt:678 Work on sheet GStruct. ----
    Call vlaensuresheet("gstruct") ' vla:1623 src:678
    Call worksheets("gstruct").activate
    vla_step = 361 ' vla:1624
    If vlatraceon() Then ' vla:1625
        Call vlatracestep(361, vla_step_text(361)) ' vla:1625
    End If
    ' ---- instructions.txt:680 Hide row 3. ----
    rows(3).hidden = True
    vla_step = 362 ' vla:1629
    If vlatraceon() Then ' vla:1630
        Call vlatracestep(362, vla_step_text(362)) ' vla:1630
    End If
    columns("b").hidden = True
    vla_step = 363 ' vla:1633
    If vlatraceon() Then ' vla:1634
        Call vlatracestep(363, vla_step_text(363)) ' vla:1634
    End If
    cells.entirerow.hidden = False
    cells.entirecolumn.hidden = False
    vla_step = 364 ' vla:1637
    If vlatraceon() Then ' vla:1638
        Call vlatracestep(364, vla_step_text(364)) ' vla:1638
    End If
    ' ---- instructions.txt:684 Set font size of cell A6 to 36. ----
    range("a6").font.size = 36
    vla_step = 365 ' vla:1642
    If vlatraceon() Then ' vla:1643
        Call vlatracestep(365, vla_step_text(365)) ' vla:1643
    End If
    Call rows(6).autofit
    vla_step = 366 ' vla:1646
    If vlatraceon() Then ' vla:1647
        Call vlatracestep(366, vla_step_text(366)) ' vla:1647
    End If
    ' ---- instructions.txt:687 Group rows 10 through 12. ----
    Call rows((10 & ":" & 12)).group
    vla_step = 367 ' vla:1651
    If vlatraceon() Then ' vla:1652
        Call vlatracestep(367, vla_step_text(367)) ' vla:1652
    End If
    Call rows((14 & ":" & 16)).group
    vla_step = 368 ' vla:1655
    If vlatraceon() Then ' vla:1656
        Call vlatracestep(368, vla_step_text(368)) ' vla:1656
    End If
    Call rows((14 & ":" & 16)).ungroup
    vla_step = 369 ' vla:1659
    If vlatraceon() Then ' vla:1660
        Call vlatracestep(369, vla_step_text(369)) ' vla:1660
    End If
    ' ---- instructions.txt:691 Put "before-insert" into cell A20. ----
    range("a20") = "before-insert" ' vla:1663 src:691
    vla_step = 370 ' vla:1664
    If vlatraceon() Then ' vla:1665
        Call vlatracestep(370, vla_step_text(370)) ' vla:1665
    End If
    Call rows(20).resize(rowsize:=3).insert
    vla_step = 371 ' vla:1668
    If vlatraceon() Then ' vla:1669
        Call vlatracestep(371, vla_step_text(371)) ' vla:1669
    End If
    ' ---- instructions.txt:694 Put "before-delete" into cell A40. ----
    range("a40") = "before-delete" ' vla:1672 src:694
    vla_step = 372 ' vla:1673
    If vlatraceon() Then ' vla:1674
        Call vlatracestep(372, vla_step_text(372)) ' vla:1674
    End If
    Call rows((38 & ":" & 39)).delete
    vla_step = 373 ' vla:1677
    If vlatraceon() Then ' vla:1678
        Call vlatracestep(373, vla_step_text(373)) ' vla:1678
    End If
    ' ---- instructions.txt:697 Put "marker-b" into cell B50. ----
    range("b50") = "marker-b" ' vla:1681 src:697
    vla_step = 374 ' vla:1682
    If vlatraceon() Then ' vla:1683
        Call vlatracestep(374, vla_step_text(374)) ' vla:1683
    End If
    range("c50") = "marker-c" ' vla:1685 src:698
    vla_step = 375 ' vla:1686
    If vlatraceon() Then ' vla:1687
        Call vlatracestep(375, vla_step_text(375)) ' vla:1687
    End If
    range("d50") = "marker-d" ' vla:1689 src:699
    vla_step = 376 ' vla:1690
    If vlatraceon() Then ' vla:1691
        Call vlatracestep(376, vla_step_text(376)) ' vla:1691
    End If
    Call vlamovecolumn("b", "d")
    vla_step = 377 ' vla:1694
    If vlatraceon() Then ' vla:1695
        Call vlatracestep(377, vla_step_text(377)) ' vla:1695
    End If
    ' ---- instructions.txt:702 Freeze the first 2 rows. ----
    Call vlafreezepanes(2)
    vla_step = 378 ' vla:1699
    If vlatraceon() Then ' vla:1700
        Call vlatracestep(378, vla_step_text(378)) ' vla:1700
    End If
    ' ---- instructions.txt:704 Make cell A60 bold. ----
    range("a60").font.bold = True
    vla_step = 379 ' vla:1704
    If vlatraceon() Then ' vla:1705
        Call vlatracestep(379, vla_step_text(379)) ' vla:1705
    End If
    range("a60").interior.color = vbred
    vla_step = 380 ' vla:1708
    If vlatraceon() Then ' vla:1709
        Call vlatracestep(380, vla_step_text(380)) ' vla:1709
    End If
    Call range("a60:a60").copy
    Call range("c60:c60").pastespecial(paste:=xlpasteformats)
    application.cutcopymode = False
    vla_step = 381 ' vla:1712
    If vlatraceon() Then ' vla:1713
        Call vlatracestep(381, vla_step_text(381)) ' vla:1713
    End If
    Call range("a60:a60").copy
    Call range("c62:c62").pastespecial(paste:=xlpasteformats)
    application.cutcopymode = False
    vla_step = 382 ' vla:1716
    If vlatraceon() Then ' vla:1717
        Call vlatracestep(382, vla_step_text(382)) ' vla:1717
    End If
    ' ---- instructions.txt:710 Put formula "=5*2" into cell B64. ----
    range("b64").Formula2 = "=5*2"
    vla_step = 383 ' vla:1721
    If vlatraceon() Then ' vla:1722
        Call vlatracestep(383, vla_step_text(383)) ' vla:1722
    End If
    Call range("b64").copy
    Call range("d64").pastespecial(paste:=xlpasteformulas)
    application.cutcopymode = False
    vla_step = 384 ' vla:1725
    If vlatraceon() Then ' vla:1726
        Call vlatracestep(384, vla_step_text(384)) ' vla:1726
    End If
    ' ---- instructions.txt:713 Set width of column F to 33. ----
    columns("f").columnwidth = 33
    vla_step = 385 ' vla:1730
    If vlatraceon() Then ' vla:1731
        Call vlatracestep(385, vla_step_text(385)) ' vla:1731
    End If
    Call range("f1:f1").copy
    Call range("h1:h1").pastespecial(paste:=xlpastecolumnwidths)
    application.cutcopymode = False
    vla_step = 386 ' vla:1734
    If vlatraceon() Then ' vla:1735
        Call vlatracestep(386, vla_step_text(386)) ' vla:1735
    End If
    ' ---- instructions.txt:716 Put 1 into cell A68. ----
    range("a68") = 1 ' vla:1738 src:716
    vla_step = 387 ' vla:1739
    If vlatraceon() Then ' vla:1740
        Call vlatracestep(387, vla_step_text(387)) ' vla:1740
    End If
    range("a69") = 2 ' vla:1742 src:717
    vla_step = 388 ' vla:1743
    If vlatraceon() Then ' vla:1744
        Call vlatracestep(388, vla_step_text(388)) ' vla:1744
    End If
    range("a70") = 3 ' vla:1746 src:718
    vla_step = 389 ' vla:1747
    If vlatraceon() Then ' vla:1748
        Call vlatracestep(389, vla_step_text(389)) ' vla:1748
    End If
    Call range("a68:a70").copy
    Call range("c68").pastespecial(transpose:=True)
    application.cutcopymode = False
    vla_step = 390 ' vla:1751
    If vlatraceon() Then ' vla:1752
        Call vlatracestep(390, vla_step_text(390)) ' vla:1752
    End If
    ' ---- instructions.txt:721 Put "cutme" into cell A72. ----
    range("a72") = "cutme" ' vla:1755 src:721
    vla_step = 391 ' vla:1756
    If vlatraceon() Then ' vla:1757
        Call vlatracestep(391, vla_step_text(391)) ' vla:1757
    End If
    Call range("a72:a72").cut(destination:=range("c72"))
    vla_step = 392 ' vla:1760
    If vlatraceon() Then ' vla:1761
        Call vlatracestep(392, vla_step_text(392)) ' vla:1761
    End If
    ' ---- instructions.txt:724 Put "rowdata" into cell A74. ----
    range("a74") = "rowdata" ' vla:1764 src:724
    vla_step = 393 ' vla:1765
    If vlatraceon() Then ' vla:1766
        Call vlatracestep(393, vla_step_text(393)) ' vla:1766
    End If
    Call rows(74).copy(destination:=rows(76))
    vla_step = 394 ' vla:1769
    If vlatraceon() Then ' vla:1770
        Call vlatracestep(394, vla_step_text(394)) ' vla:1770
    End If
    ' ---- instructions.txt:727 Put "clearme" into cell A80. ----
    range("a80") = "clearme" ' vla:1773 src:728
    vla_step = 395 ' vla:1774
    If vlatraceon() Then ' vla:1775
        Call vlatracestep(395, vla_step_text(395)) ' vla:1775
    End If
    range("a80").font.bold = True
    vla_step = 396 ' vla:1778
    If vlatraceon() Then ' vla:1779
        Call vlatracestep(396, vla_step_text(396)) ' vla:1779
    End If
    range("a80").interior.color = vbred
    vla_step = 397 ' vla:1782
    If vlatraceon() Then ' vla:1783
        Call vlatracestep(397, vla_step_text(397)) ' vla:1783
    End If
    Call range("a80:b80").clear
    vla_step = 398 ' vla:1786
    If vlatraceon() Then ' vla:1787
        Call vlatracestep(398, vla_step_text(398)) ' vla:1787
    End If
    ' ---- instructions.txt:733 Put "m1" into cell A84. ----
    range("a84") = "m1" ' vla:1790 src:733
    vla_step = 399 ' vla:1791
    If vlatraceon() Then ' vla:1792
        Call vlatracestep(399, vla_step_text(399)) ' vla:1792
    End If
    range("a85") = "m2" ' vla:1794 src:734
    vla_step = 400 ' vla:1795
    If vlatraceon() Then ' vla:1796
        Call vlatracestep(400, vla_step_text(400)) ' vla:1796
    End If
    range("a86") = "m3" ' vla:1798 src:735
    vla_step = 401 ' vla:1799
    If vlatraceon() Then ' vla:1800
        Call vlatracestep(401, vla_step_text(401)) ' vla:1800
    End If
    range("a87") = "m4" ' vla:1802 src:736
    vla_step = 402 ' vla:1803
    If vlatraceon() Then ' vla:1804
        Call vlatracestep(402, vla_step_text(402)) ' vla:1804
    End If
    range("a88") = "m5" ' vla:1806 src:737
    vla_step = 403 ' vla:1807
    If vlatraceon() Then ' vla:1808
        Call vlatracestep(403, vla_step_text(403)) ' vla:1808
    End If
    Call range("a84:a86").delete(shift:=xlshiftup)
    vla_step = 404 ' vla:1811
    If vlatraceon() Then ' vla:1812
        Call vlatracestep(404, vla_step_text(404)) ' vla:1812
    End If
    ' ---- instructions.txt:740 Put "x1" into cell B90. ----
    range("b90") = "x1" ' vla:1815 src:740
    vla_step = 405 ' vla:1816
    If vlatraceon() Then ' vla:1817
        Call vlatracestep(405, vla_step_text(405)) ' vla:1817
    End If
    range("c90") = "x2" ' vla:1819 src:741
    vla_step = 406 ' vla:1820
    If vlatraceon() Then ' vla:1821
        Call vlatracestep(406, vla_step_text(406)) ' vla:1821
    End If
    range("d90") = "x3" ' vla:1823 src:742
    vla_step = 407 ' vla:1824
    If vlatraceon() Then ' vla:1825
        Call vlatracestep(407, vla_step_text(407)) ' vla:1825
    End If
    range("e90") = "rightdata" ' vla:1827 src:743
    vla_step = 408 ' vla:1828
    If vlatraceon() Then ' vla:1829
        Call vlatracestep(408, vla_step_text(408)) ' vla:1829
    End If
    Call range("b90:d90").delete(shift:=xlshifttoleft)
    vla_step = 409 ' vla:1832
    If vlatraceon() Then ' vla:1833
        Call vlatracestep(409, vla_step_text(409)) ' vla:1833
    End If
    ' ---- instructions.txt:746 Put 1 into cell A94. ----
    range("a94") = 1 ' vla:1836 src:746
    vla_step = 410 ' vla:1837
    If vlatraceon() Then ' vla:1838
        Call vlatracestep(410, vla_step_text(410)) ' vla:1838
    End If
    range("b94") = 2 ' vla:1840 src:747
    vla_step = 411 ' vla:1841
    If vlatraceon() Then ' vla:1842
        Call vlatracestep(411, vla_step_text(411)) ' vla:1842
    End If
    range("a95") = 3 ' vla:1844 src:748
    vla_step = 412 ' vla:1845
    If vlatraceon() Then ' vla:1846
        Call vlatracestep(412, vla_step_text(412)) ' vla:1846
    End If
    range("a96") = 5 ' vla:1848 src:749
    vla_step = 413 ' vla:1849
    If vlatraceon() Then ' vla:1850
        Call vlatracestep(413, vla_step_text(413)) ' vla:1850
    End If
    range("b96") = 6 ' vla:1852 src:750
    vla_step = 414 ' vla:1853
    If vlatraceon() Then ' vla:1854
        Call vlatracestep(414, vla_step_text(414)) ' vla:1854
    End If
    Call vladeleteblankrows(range("a94:b96"))
    vla_step = 415 ' vla:1857
    If vlatraceon() Then ' vla:1858
        Call vlatracestep(415, vla_step_text(415)) ' vla:1858
    End If
    ' ---- instructions.txt:753 Put 1 into cell A150. ----
    range("a150") = 1 ' vla:1861 src:754
    vla_step = 416 ' vla:1862
    If vlatraceon() Then ' vla:1863
        Call vlatracestep(416, vla_step_text(416)) ' vla:1863
    End If
    range("a151") = 2 ' vla:1865 src:755
    vla_step = 417 ' vla:1866
    If vlatraceon() Then ' vla:1867
        Call vlatracestep(417, vla_step_text(417)) ' vla:1867
    End If
    range("a152") = 3 ' vla:1869 src:756
    vla_step = 418 ' vla:1870
    If vlatraceon() Then ' vla:1871
        Call vlatracestep(418, vla_step_text(418)) ' vla:1871
    End If
    range("a153") = 4 ' vla:1873 src:757
    vla_step = 419 ' vla:1874
    If vlatraceon() Then ' vla:1875
        Call vlatracestep(419, vla_step_text(419)) ' vla:1875
    End If
    Call vlabandrows(range("a150:d153"), vlacolor("#D9D9D9"))
    vla_step = 420 ' vla:1878
    If vlatraceon() Then ' vla:1879
        Call vlatracestep(420, vla_step_text(420)) ' vla:1879
    End If
    rows(165).font.bold = True
    rows(165).interior.color = vlacolor("#D9D9D9")
    vla_step = 421 ' vla:1882
    If vlatraceon() Then ' vla:1883
        Call vlatracestep(421, vla_step_text(421)) ' vla:1883
    End If
    ' ---- instructions.txt:761 Fill A170:A179 with a series starting at 1. ----
    Call vlafillseries(range("a170:a179"), 1, 1, "linear")
    vla_step = 422 ' vla:1887
    If vlatraceon() Then ' vla:1888
        Call vlatracestep(422, vla_step_text(422)) ' vla:1888
    End If
    Call vlafillseries(range("a180:a184"), 1, 2, "linear")
    vla_step = 423 ' vla:1891
    If vlatraceon() Then ' vla:1892
        Call vlatracestep(423, vla_step_text(423)) ' vla:1892
    End If
    Call vlafillseries(range("a190:a194"), 2, 2, "growth")
    vla_step = 424 ' vla:1895
    If vlatraceon() Then ' vla:1896
        Call vlatracestep(424, vla_step_text(424)) ' vla:1896
    End If
    Call vlafillseries(range("a200:a203"), 2, 3, "growth")
    vla_step = 425 ' vla:1899
    If vlatraceon() Then ' vla:1900
        Call vlatracestep(425, vla_step_text(425)) ' vla:1900
    End If
    ' ---- instructions.txt:767 Make cell Z500 bold. ----
    range("z500").font.bold = True
    vla_step = 426 ' vla:1904
    If vlatraceon() Then ' vla:1905
        Call vlatracestep(426, vla_step_text(426)) ' vla:1905
    End If
    Call range("z500").clear
    vla_step = 427 ' vla:1908
    If vlatraceon() Then ' vla:1909
        Call vlatracestep(427, vla_step_text(427)) ' vla:1909
    End If
    Call vlatrimsheet
    vla_step = 428 ' vla:1912
    If vlatraceon() Then ' vla:1913
        Call vlatracestep(428, vla_step_text(428)) ' vla:1913
    End If
    ' ---- instructions.txt:781 Go to sheet Output. ----
    Call worksheets("output").activate
    vla_step = 429 ' vla:1917
    If vlatraceon() Then ' vla:1918
        Call vlatracestep(429, vla_step_text(429)) ' vla:1918
    End If
    ' ---- instructions.txt:783 Tidy-up. ----
    Call tidy_up ' vla:1921 src:783
    vla_step = 430 ' vla:1922
    If vlatraceon() Then ' vla:1923
        Call vlatracestep(430, vla_step_text(430)) ' vla:1923
    End If
    application.screenupdating = True
    vla_step = 431 ' vla:1926
    If vlatraceon() Then ' vla:1927
        Call vlatracestep(431, vla_step_text(431)) ' vla:1927
    End If
    Debug.Print "report finished" ' vla:1929 src:785
    Exit Sub ' vla:1930
vla_fail: ' vla:1931
    Call vla_report_error ' vla:1932
End Sub

Public Sub vla_report_error()
    Call vlashowerror(("Something went wrong at step " & vla_step & ":" & vbcrlf & vbcrlf & vla_step_text(vla_step) & vbcrlf & vbcrlf & "Excel says: " & err.description)) ' vla:1935
End Sub

Public Function vla_step_text(ByVal n As Long) As String
    Select Case n ' vla:1940
        Case 1
            vla_step_text = "Work on sheet Output. [line 18]" ' vla:1941
            Exit Function
        Case 2
            vla_step_text = "Turn off screen updating. [line 20]" ' vla:1942
            Exit Function
        Case 3
            vla_step_text = "Create a number called total. [line 21]" ' vla:1943
            Exit Function
        Case 4
            vla_step_text = "Set total to 0. [line 22]" ' vla:1944
            Exit Function
        Case 5
            vla_step_text = "Put ""Test Report"" into cell A1. [line 23]" ' vla:1945
            Exit Function
        Case 6
            vla_step_text = "Make cell A1 bold. [line 24]" ' vla:1946
            Exit Function
        Case 7
            vla_step_text = "Set font size of cell A1 to 14. [line 25]" ' vla:1947
            Exit Function
        Case 8
            vla_step_text = "Put today into cell D1. [line 26]" ' vla:1948
            Exit Function
        Case 9
            vla_step_text = "Repeat 5 times: [line 28]" ' vla:1949
            Exit Function
        Case 10
            vla_step_text = "Increase total by counter. [line 29]" ' vla:1950
            Exit Function
        Case 11
            vla_step_text = "If counter is divisible by 2, log ""even step "" joined with counter. [line 30]" ' vla:1951
            Exit Function
        Case 12
            vla_step_text = "Put total into cell B2. [line 32]" ' vla:1952
            Exit Function
        Case 13
            vla_step_text = "Put formula ""=B2*2"" into cell B3. [line 33]" ' vla:1953
            Exit Function
        Case 14
            vla_step_text = "Put sum of range ""B2:B3"" into cell B4. [line 37]" ' vla:1954
            Exit Function
        Case 15
            vla_step_text = "Set grand to sum of range B2:B3. [line 38]" ' vla:1955
            Exit Function
        Case 16
            vla_step_text = "Log ""grand is "" joined with grand. [line 39]" ' vla:1956
            Exit Function
        Case 17
            vla_step_text = "If grand is greater than 40: [line 43]" ' vla:1957
            Exit Function
        Case 18
            vla_step_text = "Put ""PASS"" into cell C4. [line 44]" ' vla:1958
            Exit Function
        Case 19
            vla_step_text = "Repeat 2 times: [line 45]" ' vla:1959
            Exit Function
        Case 20
            vla_step_text = "Log ""pass check "" joined with counter. [line 46]" ' vla:1960
            Exit Function
        Case 21
            vla_step_text = "Make cell C4 bold. [line 48]" ' vla:1961
            Exit Function
        Case 22
            vla_step_text = "Put ""CHECK"" into cell C4. [line 51]" ' vla:1962
            Exit Function
        Case 23
            vla_step_text = "If grand is at least 45, make cell C4 yellow. [line 53]" ' vla:1963
            Exit Function
        Case 24
            vla_step_text = "Create a text called label. [line 54]" ' vla:1964
            Exit Function
        Case 25
            vla_step_text = "Set label to ""Total: "" joined with total. [line 55]" ' vla:1965
            Exit Function
        Case 26
            vla_step_text = "Put label into cell A6. [line 56]" ' vla:1966
            Exit Function
        Case 27
            vla_step_text = "Set font-color of cell A6 to hot-pink. [line 57]" ' vla:1967
            Exit Function
        Case 28
            vla_step_text = "Remember range B2:B4 as results. [line 58]" ' vla:1968
            Exit Function
        Case 29
            vla_step_text = "For each r in results, log r. [line 59]" ' vla:1969
            Exit Function
        Case 30
            vla_step_text = "Set biggest to largest of results. [line 60]" ' vla:1970
            Exit Function
        Case 31
            vla_step_text = "Log ""largest result is "" joined with biggest joined with "", label length "" joined with length of label. [line 61]" ' vla:1971
            Exit Function
        Case 32
            vla_step_text = "Repeat 3 times: [line 64]" ' vla:1972
            Exit Function
        Case 33
            vla_step_text = "Put counter times 10 into column E row counter. [line 65]" ' vla:1973
            Exit Function
        Case 34
            vla_step_text = "Set probe to cell in column E row 2. [line 67]" ' vla:1974
            Exit Function
        Case 35
            vla_step_text = "Create a text called num-col-check. [line 68]" ' vla:1975
            Exit Function
        Case 36
            vla_step_text = "If cell in column number 5 row 2 is 20, set num-col-check to ""yes"". [line 69]" ' vla:1976
            Exit Function
        Case 37
            vla_step_text = "Create a text called value-word-check. [line 70]" ' vla:1977
            Exit Function
        Case 38
            vla_step_text = "If value in column number 5 row 3 is 30, set value-word-check to ""ok"". [line 71]" ' vla:1978
            Exit Function
        Case 39
            vla_step_text = "Put ""bang"" into cell Output!H16. [line 72]" ' vla:1979
            Exit Function
        Case 40
            vla_step_text = "Center cell A1. [line 73]" ' vla:1980
            Exit Function
        Case 41
            vla_step_text = "Set width of column D to 24. [line 74]" ' vla:1981
            Exit Function
        Case 42
            vla_step_text = "Set round-check to 3.14159 rounded to 2 decimals. [line 75]" ' vla:1982
            Exit Function
        Case 43
            vla_step_text = "Log ""rounded is "" joined with round-check. [line 76]" ' vla:1983
            Exit Function
        Case 44
            vla_step_text = "Set thousand-check to 1,000 plus 500. [line 77]" ' vla:1984
            Exit Function
        Case 45
            vla_step_text = "Log ""thousands read as "" joined with thousand-check. [line 78]" ' vla:1985
            Exit Function
        Case 46
            vla_step_text = "Log ""probe is "" joined with probe. [line 79]" ' vla:1986
            Exit Function
        Case 47
            vla_step_text = "Create a number called countdown. [line 81]" ' vla:1987
            Exit Function
        Case 48
            vla_step_text = "Set countdown to 3. [line 82]" ' vla:1988
            Exit Function
        Case 49
            vla_step_text = "While countdown is greater than 0: [line 83]" ' vla:1989
            Exit Function
        Case 50
            vla_step_text = "Log ""countdown "" joined with countdown. [line 84]" ' vla:1990
            Exit Function
        Case 51
            vla_step_text = "Decrease countdown by 1. [line 85]" ' vla:1991
            Exit Function
        Case 52
            vla_step_text = "Put value into column F row row-number. [line 90]" ' vla:1992
            Exit Function
        Case 53
            vla_step_text = "Stamp. [line 92]" ' vla:1993
            Exit Function
        Case 54
            vla_step_text = "Stamp with row-number of 2 and value of ""beta"". [line 93]" ' vla:1994
            Exit Function
        Case 55
            vla_step_text = "Set f-last to last filled row of column F. [line 94]" ' vla:1995
            Exit Function
        Case 56
            vla_step_text = "Log ""column F filled to row "" joined with f-last. [line 95]" ' vla:1996
            Exit Function
        Case 57
            vla_step_text = "Count echo-row from 1 to f-last, log ""echo "" joined with echo-row. [line 96]" ' vla:1997
            Exit Function
        Case 58
            vla_step_text = "Count check-row from 1 to f-last: [line 97]" ' vla:1998
            Exit Function
        Case 59
            vla_step_text = "If cell in column F row check-row contains ""ok"", make cell in column F row check-row bold. [line 98]" ' vla:1999
            Exit Function
        Case 60
            vla_step_text = "Count stripe-row from 1 to f-last step 2: [line 103]" ' vla:2000
            Exit Function
        Case 61
            vla_step_text = "Make cell in column F row stripe-row bold. [line 104]" ' vla:2001
            Exit Function
        Case 62
            vla_step_text = "Count back-row down from f-last to 1 step 2: [line 106]" ' vla:2002
            Exit Function
        Case 63
            vla_step_text = "Log ""back-row "" joined with back-row. [line 107]" ' vla:2003
            Exit Function
        Case 64
            vla_step_text = "Count search-row from 1 to 10: [line 111]" ' vla:2004
            Exit Function
        Case 65
            vla_step_text = "If search-row is 3, stop the loop. [line 112]" ' vla:2005
            Exit Function
        Case 66
            vla_step_text = "Log ""stopped at "" joined with search-row. [line 114]" ' vla:2006
            Exit Function
        Case 67
            vla_step_text = "Create a list called found-items. [line 119]" ' vla:2007
            Exit Function
        Case 68
            vla_step_text = "Repeat 4 times: [line 120]" ' vla:2008
            Exit Function
        Case 69
            vla_step_text = "If counter is greater than 2, append counter times 100 to found-items. [line 121]" ' vla:2009
            Exit Function
        Case 70
            vla_step_text = "For each f in found-items, log ""found "" joined with f. [line 123]" ' vla:2010
            Exit Function
        Case 71
            vla_step_text = "Set list-count to count of found-items. [line 124]" ' vla:2011
            Exit Function
        Case 72
            vla_step_text = "Log ""list holds "" joined with list-count. [line 125]" ' vla:2012
            Exit Function
        Case 73
            vla_step_text = "Create a text called verdict. [line 131]" ' vla:2013
            Exit Function
        Case 74
            vla_step_text = "If grand is greater than 100: [line 132]" ' vla:2014
            Exit Function
        Case 75
            vla_step_text = "Set verdict to ""huge"". [line 133]" ' vla:2015
            Exit Function
        Case 76
            vla_step_text = "Set verdict to ""solid"". [line 136]" ' vla:2016
            Exit Function
        Case 77
            vla_step_text = "Set verdict to ""small"". [line 139]" ' vla:2017
            Exit Function
        Case 78
            vla_step_text = "Create a text called region-label. [line 141]" ' vla:2018
            Exit Function
        Case 79
            vla_step_text = "Set region to ""South"". [line 142]" ' vla:2019
            Exit Function
        Case 80
            vla_step_text = "When region is ""North"": [line 143]" ' vla:2020
            Exit Function
        Case 81
            vla_step_text = "Set region-label to ""cold"". [line 144]" ' vla:2021
            Exit Function
        Case 82
            vla_step_text = "Set region-label to ""warm"". [line 147]" ' vla:2022
            Exit Function
        Case 83
            vla_step_text = "Set region-label to ""unknown"". [line 150]" ' vla:2023
            Exit Function
        Case 84
            vla_step_text = "Log ""verdict "" joined with verdict joined with "", region "" joined with region-label. [line 152]" ' vla:2024
            Exit Function
        Case 85
            vla_step_text = "Create a number called until-count. [line 154]" ' vla:2025
            Exit Function
        Case 86
            vla_step_text = "Set fuel to 3. [line 155]" ' vla:2026
            Exit Function
        Case 87
            vla_step_text = "Repeat until fuel is 0: [line 156]" ' vla:2027
            Exit Function
        Case 88
            vla_step_text = "Decrease fuel by 1. [line 157]" ' vla:2028
            Exit Function
        Case 89
            vla_step_text = "Increase until-count by 1. [line 158]" ' vla:2029
            Exit Function
        Case 90
            vla_step_text = "Log ""repeat-until ran "" joined with until-count joined with "" times"". [line 160]" ' vla:2030
            Exit Function
        Case 91
            vla_step_text = "Create a text called rescue. [line 166]" ' vla:2031
            Exit Function
        Case 92
            vla_step_text = "Try: [line 167]" ' vla:2032
            Exit Function
        Case 93
            vla_step_text = "Go to sheet Nowhere-Land. [line 168]" ' vla:2033
            Exit Function
        Case 94
            vla_step_text = "Set rescue to ""unreachable"". [line 169]" ' vla:2034
            Exit Function
        Case 95
            vla_step_text = "Log ""the problem was "" joined with the problem. [line 172]" ' vla:2035
            Exit Function
        Case 96
            vla_step_text = "If the problem is not empty, set rescue to ""rescued"". [line 173]" ' vla:2036
            Exit Function
        Case 97
            vla_step_text = "Create a number called risk-free. [line 175]" ' vla:2037
            Exit Function
        Case 98
            vla_step_text = "Try: [line 176]" ' vla:2038
            Exit Function
        Case 99
            vla_step_text = "Set risk-free to 7. [line 177]" ' vla:2039
            Exit Function
        Case 100
            vla_step_text = "Set risk-free to -1. [line 180]" ' vla:2040
            Exit Function
        Case 101
            vla_step_text = "Log ""rescue "" joined with rescue joined with "", risk-free "" joined with risk-free. [line 182]" ' vla:2041
            Exit Function
        Case 102
            vla_step_text = "Give back amount times 0.08. [line 187]" ' vla:2042
            Exit Function
        Case 103
            vla_step_text = "Set fee to tax of 100. [line 189]" ' vla:2043
            Exit Function
        Case 104
            vla_step_text = "Log ""fee is "" joined with fee. [line 190]" ' vla:2044
            Exit Function
        Case 105
            vla_step_text = "Create a text called fee-size. [line 191]" ' vla:2045
            Exit Function
        Case 106
            vla_step_text = "If tax of 50 is greater than 3: [line 192]" ' vla:2046
            Exit Function
        Case 107
            vla_step_text = "Set fee-size to ""big"". [line 193]" ' vla:2047
            Exit Function
        Case 108
            vla_step_text = "Set fee-size to ""small"". [line 196]" ' vla:2048
            Exit Function
        Case 109
            vla_step_text = "Fit column A. [line 199]" ' vla:2049
            Exit Function
        Case 110
            vla_step_text = "Fit column B. [line 200]" ' vla:2050
            Exit Function
        Case 111
            vla_step_text = "Set font-color of cell A1 to hot-pink. [line 201]" ' vla:2051
            Exit Function
        Case 112
            vla_step_text = "Give back sale times rate. [line 206]" ' vla:2052
            Exit Function
        Case 113
            vla_step_text = "Set full-commission to commission using sale of 2000 and rate of 10%. [line 208]" ' vla:2053
            Exit Function
        Case 114
            vla_step_text = "Set default-commission to get commission using sale of 600. [line 209]" ' vla:2054
            Exit Function
        Case 115
            vla_step_text = "Log ""commissions "" joined with full-commission joined with "" / "" joined with default-commission. [line 210]" ' vla:2055
            Exit Function
        Case 116
            vla_step_text = "Try: [line 219]" ' vla:2056
            Exit Function
        Case 117
            vla_step_text = "Add sheet called ""Q1 Data"". [line 220]" ' vla:2057
            Exit Function
        Case 118
            vla_step_text = "Go to sheet Output. [line 222]" ' vla:2058
            Exit Function
        Case 119
            vla_step_text = "Put ""spaced"" into cell 'Q1 Data'!A1. [line 223]" ' vla:2059
            Exit Function
        Case 120
            vla_step_text = "Create a text called spaced-check. [line 224]" ' vla:2060
            Exit Function
        Case 121
            vla_step_text = "Set spaced-check to value in cell 'Q1 Data'!A1. [line 225]" ' vla:2061
            Exit Function
        Case 122
            vla_step_text = "Give back 20%. [line 230]" ' vla:2062
            Exit Function
        Case 123
            vla_step_text = "Create a number called growth-check. [line 237]" ' vla:2063
            Exit Function
        Case 124
            vla_step_text = "Set growth-check to 200. [line 238]" ' vla:2064
            Exit Function
        Case 125
            vla_step_text = "Grow growth-check by 10%. [line 239]" ' vla:2065
            Exit Function
        Case 126
            vla_step_text = "Increase growth-check by 50%. [line 240]" ' vla:2066
            Exit Function
        Case 127
            vla_step_text = "Add 100 percent to growth-check. [line 241]" ' vla:2067
            Exit Function
        Case 128
            vla_step_text = "Decrease growth-check by 75%. [line 242]" ' vla:2068
            Exit Function
        Case 129
            vla_step_text = "Set pick-check to item 2 of found-items. [line 243]" ' vla:2069
            Exit Function
        Case 130
            vla_step_text = "Log ""grew to "" joined with growth-check joined with "", picked "" joined with pick-check joined with "", first "" joined with first of found-items. [line 244]" ' vla:2070
            Exit Function
        Case 131
            vla_step_text = "Create a text called quote-check. [line 245]" ' vla:2071
            Exit Function
        Case 132
            vla_step_text = "Set quote-check to ""He said """"ok"""""". [line 246]" ' vla:2072
            Exit Function
        Case 133
            vla_step_text = "Create a lookup called prices. [line 260]" ' vla:2073
            Exit Function
        Case 134
            vla_step_text = "Store 100 at key ""ax-7"" in prices. [line 261]" ' vla:2074
            Exit Function
        Case 135
            vla_step_text = "Store 250 under ""bx-2"" in prices. [line 262]" ' vla:2075
            Exit Function
        Case 136
            vla_step_text = "Store 120 at key ""AX-7"" in prices. [line 263]" ' vla:2076
            Exit Function
        Case 137
            vla_step_text = "Set ax-price to prices for ""ax-7"". [line 264]" ' vla:2077
            Exit Function
        Case 138
            vla_step_text = "Set key-count to count of keys of prices. [line 265]" ' vla:2078
            Exit Function
        Case 139
            vla_step_text = "Set price-sum to prices for ""ax-7"" plus prices for ""bx-2"". [line 266]" ' vla:2079
            Exit Function
        Case 140
            vla_step_text = "Create a text called key-list. [line 267]" ' vla:2080
            Exit Function
        Case 141
            vla_step_text = "For each k in keys of prices: [line 268]" ' vla:2081
            Exit Function
        Case 142
            vla_step_text = "Set key-list to key-list joined with k. [line 269]" ' vla:2082
            Exit Function
        Case 143
            vla_step_text = "Create a text called price-verdict. [line 271]" ' vla:2083
            Exit Function
        Case 144
            vla_step_text = "If prices for ""bx-2"" is greater than 200, set price-verdict to ""steep"". [line 272]" ' vla:2084
            Exit Function
        Case 145
            vla_step_text = "Create a text called pair-trace. [line 273]" ' vla:2085
            Exit Function
        Case 146
            vla_step_text = "For each pair in prices: [line 274]" ' vla:2086
            Exit Function
        Case 147
            vla_step_text = "If value of pair is greater than 200, set pair-trace to pair-trace joined with key of pair. [line 275]" ' vla:2087
            Exit Function
        Case 148
            vla_step_text = "Log ""lookup: ax "" joined with ax-price joined with "", keys "" joined with key-list joined with "", pairs "" joined with pair-trace. [line 277]" ' vla:2088
            Exit Function
        Case 149
            vla_step_text = "Put ""Item"" into cell J1. [line 284]" ' vla:2089
            Exit Function
        Case 150
            vla_step_text = "Put ""Amount"" into cell K1. [line 285]" ' vla:2090
            Exit Function
        Case 151
            vla_step_text = "Put ""Widget"" into cell J2. [line 286]" ' vla:2091
            Exit Function
        Case 152
            vla_step_text = "Put 10 into cell K2. [line 287]" ' vla:2092
            Exit Function
        Case 153
            vla_step_text = "Put ""Gadget"" into cell J3. [line 288]" ' vla:2093
            Exit Function
        Case 154
            vla_step_text = "Put 20 into cell K3. [line 289]" ' vla:2094
            Exit Function
        Case 155
            vla_step_text = "Turn J1:K3 into a table called SalesTable. [line 290]" ' vla:2095
            Exit Function
        Case 156
            vla_step_text = "Set style of table SalesTable to TableStyleMedium9. [line 291]" ' vla:2096
            Exit Function
        Case 157
            vla_step_text = "Show the total row of table SalesTable. [line 292]" ' vla:2097
            Exit Function
        Case 158
            vla_step_text = "Put ""X"" into cell J5. [line 296]" ' vla:2098
            Exit Function
        Case 159
            vla_step_text = "Put ""Y"" into cell J6. [line 297]" ' vla:2099
            Exit Function
        Case 160
            vla_step_text = "Turn J5:J6 into a table called QuietTable. [line 298]" ' vla:2100
            Exit Function
        Case 161
            vla_step_text = "Show the total row of table QuietTable. [line 299]" ' vla:2101
            Exit Function
        Case 162
            vla_step_text = "Hide the total row of table QuietTable. [line 300]" ' vla:2102
            Exit Function
        Case 163
            vla_step_text = "Put ""A"" into cell J8. [line 304]" ' vla:2103
            Exit Function
        Case 164
            vla_step_text = "Put ""B"" into cell J9. [line 305]" ' vla:2104
            Exit Function
        Case 165
            vla_step_text = "Turn J8:J9 into a table called TempTable. [line 306]" ' vla:2105
            Exit Function
        Case 166
            vla_step_text = "Turn table TempTable back into a range. [line 307]" ' vla:2106
            Exit Function
        Case 167
            vla_step_text = "Put ""Item"" into cell J11. [line 312]" ' vla:2107
            Exit Function
        Case 168
            vla_step_text = "Put ""Qty"" into cell K11. [line 313]" ' vla:2108
            Exit Function
        Case 169
            vla_step_text = "Put ""Bolt"" into cell J12. [line 314]" ' vla:2109
            Exit Function
        Case 170
            vla_step_text = "Put 5 into cell K12. [line 315]" ' vla:2110
            Exit Function
        Case 171
            vla_step_text = "Put ""Nut"" into cell J13. [line 316]" ' vla:2111
            Exit Function
        Case 172
            vla_step_text = "Put 8 into cell K13. [line 317]" ' vla:2112
            Exit Function
        Case 173
            vla_step_text = "Turn J11:K13 into a table called EditTable. [line 318]" ' vla:2113
            Exit Function
        Case 174
            vla_step_text = "Add a row to table EditTable. [line 319]" ' vla:2114
            Exit Function
        Case 175
            vla_step_text = "Delete row 1 of table EditTable. [line 320]" ' vla:2115
            Exit Function
        Case 176
            vla_step_text = "Set qty-values to column Qty of table EditTable. [line 321]" ' vla:2116
            Exit Function
        Case 177
            vla_step_text = "For each q in qty-values, log q. [line 322]" ' vla:2117
            Exit Function
        Case 178
            vla_step_text = "Put ""Region"" into cell N1. [line 332]" ' vla:2118
            Exit Function
        Case 179
            vla_step_text = "Put ""Product"" into cell O1. [line 333]" ' vla:2119
            Exit Function
        Case 180
            vla_step_text = "Put ""Segment"" into cell P1. [line 334]" ' vla:2120
            Exit Function
        Case 181
            vla_step_text = "Put ""Channel"" into cell Q1. [line 335]" ' vla:2121
            Exit Function
        Case 182
            vla_step_text = "Put ""Units"" into cell R1. [line 336]" ' vla:2122
            Exit Function
        Case 183
            vla_step_text = "Put ""Revenue"" into cell S1. [line 337]" ' vla:2123
            Exit Function
        Case 184
            vla_step_text = "Put ""North"" into cell N2. [line 338]" ' vla:2124
            Exit Function
        Case 185
            vla_step_text = "Put ""Widget"" into cell O2. [line 339]" ' vla:2125
            Exit Function
        Case 186
            vla_step_text = "Put ""Retail"" into cell P2. [line 340]" ' vla:2126
            Exit Function
        Case 187
            vla_step_text = "Put ""Online"" into cell Q2. [line 341]" ' vla:2127
            Exit Function
        Case 188
            vla_step_text = "Put 10 into cell R2. [line 342]" ' vla:2128
            Exit Function
        Case 189
            vla_step_text = "Put 500 into cell S2. [line 343]" ' vla:2129
            Exit Function
        Case 190
            vla_step_text = "Put ""North"" into cell N3. [line 344]" ' vla:2130
            Exit Function
        Case 191
            vla_step_text = "Put ""Gadget"" into cell O3. [line 345]" ' vla:2131
            Exit Function
        Case 192
            vla_step_text = "Put ""Wholesale"" into cell P3. [line 346]" ' vla:2132
            Exit Function
        Case 193
            vla_step_text = "Put ""Store"" into cell Q3. [line 347]" ' vla:2133
            Exit Function
        Case 194
            vla_step_text = "Put 5 into cell R3. [line 348]" ' vla:2134
            Exit Function
        Case 195
            vla_step_text = "Put 200 into cell S3. [line 349]" ' vla:2135
            Exit Function
        Case 196
            vla_step_text = "Put ""South"" into cell N4. [line 350]" ' vla:2136
            Exit Function
        Case 197
            vla_step_text = "Put ""Widget"" into cell O4. [line 351]" ' vla:2137
            Exit Function
        Case 198
            vla_step_text = "Put ""Wholesale"" into cell P4. [line 352]" ' vla:2138
            Exit Function
        Case 199
            vla_step_text = "Put ""Online"" into cell Q4. [line 353]" ' vla:2139
            Exit Function
        Case 200
            vla_step_text = "Put 20 into cell R4. [line 354]" ' vla:2140
            Exit Function
        Case 201
            vla_step_text = "Put 900 into cell S4. [line 355]" ' vla:2141
            Exit Function
        Case 202
            vla_step_text = "Put ""South"" into cell N5. [line 356]" ' vla:2142
            Exit Function
        Case 203
            vla_step_text = "Put ""Gadget"" into cell O5. [line 357]" ' vla:2143
            Exit Function
        Case 204
            vla_step_text = "Put ""Retail"" into cell P5. [line 358]" ' vla:2144
            Exit Function
        Case 205
            vla_step_text = "Put ""Store"" into cell Q5. [line 359]" ' vla:2145
            Exit Function
        Case 206
            vla_step_text = "Put 8 into cell R5. [line 360]" ' vla:2146
            Exit Function
        Case 207
            vla_step_text = "Put 300 into cell S5. [line 361]" ' vla:2147
            Exit Function
        Case 208
            vla_step_text = "Put ""East"" into cell N6. [line 362]" ' vla:2148
            Exit Function
        Case 209
            vla_step_text = "Put ""Widget"" into cell O6. [line 363]" ' vla:2149
            Exit Function
        Case 210
            vla_step_text = "Put ""Retail"" into cell P6. [line 364]" ' vla:2150
            Exit Function
        Case 211
            vla_step_text = "Put ""Online"" into cell Q6. [line 365]" ' vla:2151
            Exit Function
        Case 212
            vla_step_text = "Put 12 into cell R6. [line 366]" ' vla:2152
            Exit Function
        Case 213
            vla_step_text = "Put 600 into cell S6. [line 367]" ' vla:2153
            Exit Function
        Case 214
            vla_step_text = "Put ""East"" into cell N7. [line 368]" ' vla:2154
            Exit Function
        Case 215
            vla_step_text = "Put ""Gadget"" into cell O7. [line 369]" ' vla:2155
            Exit Function
        Case 216
            vla_step_text = "Put ""Wholesale"" into cell P7. [line 370]" ' vla:2156
            Exit Function
        Case 217
            vla_step_text = "Put ""Store"" into cell Q7. [line 371]" ' vla:2157
            Exit Function
        Case 218
            vla_step_text = "Put 6 into cell R7. [line 372]" ' vla:2158
            Exit Function
        Case 219
            vla_step_text = "Put 250 into cell S7. [line 373]" ' vla:2159
            Exit Function
        Case 220
            vla_step_text = "Make a pivot table from N1:S7 at U1 called SalesPivot. [line 381]" ' vla:2160
            Exit Function
        Case 221
            vla_step_text = "Add rows of Region, Product to pivot SalesPivot. [line 382]" ' vla:2161
            Exit Function
        Case 222
            vla_step_text = "Add columns of Segment to pivot SalesPivot. [line 383]" ' vla:2162
            Exit Function
        Case 223
            vla_step_text = "Add filters of Channel to pivot SalesPivot. [line 384]" ' vla:2163
            Exit Function
        Case 224
            vla_step_text = "Add Revenue to pivot SalesPivot as a sum. [line 392]" ' vla:2164
            Exit Function
        Case 225
            vla_step_text = "Add Revenue, Units to pivot SalesPivot as a count. [line 393]" ' vla:2165
            Exit Function
        Case 226
            vla_step_text = "Add Units to pivot SalesPivot as an average. [line 394]" ' vla:2166
            Exit Function
        Case 227
            vla_step_text = "Make a pivot table from N1:S7 at N20 called FullPivot with rows of Region, Product, and Channel and columns of Segment and values of Revenue, Units. [line 406]" ' vla:2167
            Exit Function
        Case 228
            vla_step_text = "Refresh pivot SalesPivot. [line 415]" ' vla:2168
            Exit Function
        Case 229
            vla_step_text = "Refresh every pivot table. [line 416]" ' vla:2169
            Exit Function
        Case 230
            vla_step_text = "Collapse Region in pivot SalesPivot. [line 417]" ' vla:2170
            Exit Function
        Case 231
            vla_step_text = "Collapse Product in pivot FullPivot. [line 427]" ' vla:2171
            Exit Function
        Case 232
            vla_step_text = "Expand Product in pivot FullPivot. [line 428]" ' vla:2172
            Exit Function
        Case 233
            vla_step_text = "Show pivot SalesPivot in tabular form. [line 443]" ' vla:2173
            Exit Function
        Case 234
            vla_step_text = "Show pivot SalesPivot in compact form. [line 444]" ' vla:2174
            Exit Function
        Case 235
            vla_step_text = "Show pivot FullPivot in outline form. [line 445]" ' vla:2175
            Exit Function
        Case 236
            vla_step_text = "Hide subtotals for Region in pivot SalesPivot. [line 453]" ' vla:2176
            Exit Function
        Case 237
            vla_step_text = "Hide subtotals for Product in pivot FullPivot. [line 454]" ' vla:2177
            Exit Function
        Case 238
            vla_step_text = "Show subtotals for Product in pivot FullPivot. [line 455]" ' vla:2178
            Exit Function
        Case 239
            vla_step_text = "Add a blank row after Region in pivot SalesPivot. [line 457]" ' vla:2179
            Exit Function
        Case 240
            vla_step_text = "Add a blank row after Product in pivot FullPivot. [line 458]" ' vla:2180
            Exit Function
        Case 241
            vla_step_text = "Remove a blank row after Product in pivot FullPivot. [line 459]" ' vla:2181
            Exit Function
        Case 242
            vla_step_text = "Sort Region in pivot SalesPivot descending. [line 476]" ' vla:2182
            Exit Function
        Case 243
            vla_step_text = "Sort Region in pivot SalesPivot ascending. [line 477]" ' vla:2183
            Exit Function
        Case 244
            vla_step_text = "Sort Product in pivot FullPivot descending by Revenue. [line 478]" ' vla:2184
            Exit Function
        Case 245
            vla_step_text = "Sort Product in pivot FullPivot ascending by Revenue. [line 479]" ' vla:2185
            Exit Function
        Case 246
            vla_step_text = "Make a pivot table from N1:S7 at N200 called RenameMePivot. [line 490]" ' vla:2186
            Exit Function
        Case 247
            vla_step_text = "Rename pivot RenameMePivot to RenamedPivot. [line 491]" ' vla:2187
            Exit Function
        Case 248
            vla_step_text = "Make a pivot table from N1:S7 at N220 called ClearMePivot. [line 493]" ' vla:2188
            Exit Function
        Case 249
            vla_step_text = "Add rows of Region to pivot ClearMePivot. [line 494]" ' vla:2189
            Exit Function
        Case 250
            vla_step_text = "Add Revenue to pivot ClearMePivot as a sum. [line 495]" ' vla:2190
            Exit Function
        Case 251
            vla_step_text = "Clear pivot ClearMePivot. [line 496]" ' vla:2191
            Exit Function
        Case 252
            vla_step_text = "Make a pivot table from N1:S7 at N240 called RemoveFieldMePivot. [line 498]" ' vla:2192
            Exit Function
        Case 253
            vla_step_text = "Add rows of Region, Product to pivot RemoveFieldMePivot. [line 499]" ' vla:2193
            Exit Function
        Case 254
            vla_step_text = "Remove Product from pivot RemoveFieldMePivot. [line 500]" ' vla:2194
            Exit Function
        Case 255
            vla_step_text = "Put ""West"" into cell N8. [line 506]" ' vla:2195
            Exit Function
        Case 256
            vla_step_text = "Put ""Widget"" into cell O8. [line 507]" ' vla:2196
            Exit Function
        Case 257
            vla_step_text = "Put ""Retail"" into cell P8. [line 508]" ' vla:2197
            Exit Function
        Case 258
            vla_step_text = "Put ""Online"" into cell Q8. [line 509]" ' vla:2198
            Exit Function
        Case 259
            vla_step_text = "Put 15 into cell R8. [line 510]" ' vla:2199
            Exit Function
        Case 260
            vla_step_text = "Put 700 into cell S8. [line 511]" ' vla:2200
            Exit Function
        Case 261
            vla_step_text = "Make a pivot table from N1:S7 at N260 called SourceTestPivot. [line 513]" ' vla:2201
            Exit Function
        Case 262
            vla_step_text = "Add rows of Region to pivot SourceTestPivot. [line 514]" ' vla:2202
            Exit Function
        Case 263
            vla_step_text = "Change the source of pivot SourceTestPivot to N1:S8. [line 515]" ' vla:2203
            Exit Function
        Case 264
            vla_step_text = "Make a pivot table from N1:S7 at N40 called TempPivot. [line 520]" ' vla:2204
            Exit Function
        Case 265
            vla_step_text = "Delete pivot TempPivot. [line 521]" ' vla:2205
            Exit Function
        Case 266
            vla_step_text = "Put grand into cell H1. [line 524]" ' vla:2206
            Exit Function
        Case 267
            vla_step_text = "Put biggest into cell H2. [line 525]" ' vla:2207
            Exit Function
        Case 268
            vla_step_text = "Put round-check into cell H3. [line 526]" ' vla:2208
            Exit Function
        Case 269
            vla_step_text = "Put thousand-check into cell H4. [line 527]" ' vla:2209
            Exit Function
        Case 270
            vla_step_text = "Put search-row into cell H5. [line 528]" ' vla:2210
            Exit Function
        Case 271
            vla_step_text = "Put f-last into cell H6. [line 529]" ' vla:2211
            Exit Function
        Case 272
            vla_step_text = "Put list-count into cell H7. [line 530]" ' vla:2212
            Exit Function
        Case 273
            vla_step_text = "Put verdict into cell H8. [line 531]" ' vla:2213
            Exit Function
        Case 274
            vla_step_text = "Put region-label into cell H9. [line 532]" ' vla:2214
            Exit Function
        Case 275
            vla_step_text = "Put until-count into cell H10. [line 533]" ' vla:2215
            Exit Function
        Case 276
            vla_step_text = "Put rescue into cell H11. [line 534]" ' vla:2216
            Exit Function
        Case 277
            vla_step_text = "Put risk-free into cell H12. [line 535]" ' vla:2217
            Exit Function
        Case 278
            vla_step_text = "Put fee into cell H13. [line 536]" ' vla:2218
            Exit Function
        Case 279
            vla_step_text = "Put fee-size into cell H14. [line 537]" ' vla:2219
            Exit Function
        Case 280
            vla_step_text = "Put num-col-check into cell H15. [line 538]" ' vla:2220
            Exit Function
        Case 281
            vla_step_text = "Put value-word-check into cell H17. [line 539]" ' vla:2221
            Exit Function
        Case 282
            vla_step_text = "Put spaced-check into cell H18. [line 540]" ' vla:2222
            Exit Function
        Case 283
            vla_step_text = "Put full-commission into cell H19. [line 541]" ' vla:2223
            Exit Function
        Case 284
            vla_step_text = "Put default-commission into cell H20. [line 542]" ' vla:2224
            Exit Function
        Case 285
            vla_step_text = "Put growth-check into cell H21. [line 543]" ' vla:2225
            Exit Function
        Case 286
            vla_step_text = "Put pick-check into cell H22. [line 544]" ' vla:2226
            Exit Function
        Case 287
            vla_step_text = "Put quote-check into cell H23. [line 545]" ' vla:2227
            Exit Function
        Case 288
            vla_step_text = "Put vat-rate into cell H24. [line 546]" ' vla:2228
            Exit Function
        Case 289
            vla_step_text = "Put ax-price into cell H26. [line 547]" ' vla:2229
            Exit Function
        Case 290
            vla_step_text = "Put key-count into cell H27. [line 548]" ' vla:2230
            Exit Function
        Case 291
            vla_step_text = "Put key-list into cell H28. [line 549]" ' vla:2231
            Exit Function
        Case 292
            vla_step_text = "Put price-sum into cell H29. [line 550]" ' vla:2232
            Exit Function
        Case 293
            vla_step_text = "Put price-verdict into cell H30. [line 551]" ' vla:2233
            Exit Function
        Case 294
            vla_step_text = "Put pair-trace into cell H31. [line 552]" ' vla:2234
            Exit Function
        Case 295
            vla_step_text = "(set! (range ""h25"") ""vla-row"") [line 559]" ' vla:2235
            Exit Function
        Case 296
            vla_step_text = "Repeat 2 times: [line 560]" ' vla:2236
            Exit Function
        Case 297
            vla_step_text = "(debug-print counter) [line 561]" ' vla:2237
            Exit Function
        Case 298
            vla_step_text = "Work on sheet Demo. [line 576]" ' vla:2238
            Exit Function
        Case 299
            vla_step_text = "Make cell A1 italic. [line 579]" ' vla:2239
            Exit Function
        Case 300
            vla_step_text = "Make cell A2 red. [line 580]" ' vla:2240
            Exit Function
        Case 301
            vla_step_text = "Make cell A3 yellow. [line 581]" ' vla:2241
            Exit Function
        Case 302
            vla_step_text = "Set font-color of cell A4 to hot-pink. [line 582]" ' vla:2242
            Exit Function
        Case 303
            vla_step_text = "Set fill-color of cell A5 to ""#FF69B4"". [line 583]" ' vla:2243
            Exit Function
        Case 304
            vla_step_text = "Clear color of cell A5. [line 584]" ' vla:2244
            Exit Function
        Case 305
            vla_step_text = "Add border to range A1:E10. [line 585]" ' vla:2245
            Exit Function
        Case 306
            vla_step_text = "Format cell B1 as currency. [line 586]" ' vla:2246
            Exit Function
        Case 307
            vla_step_text = "Format cell B2 as percent. [line 587]" ' vla:2247
            Exit Function
        Case 308
            vla_step_text = "Format cell B3 as date. [line 588]" ' vla:2248
            Exit Function
        Case 309
            vla_step_text = "Center cell A1. [line 589]" ' vla:2249
            Exit Function
        Case 310
            vla_step_text = "Center range A1:E1. [line 590]" ' vla:2250
            Exit Function
        Case 311
            vla_step_text = "Align range A2:A5 left. [line 591]" ' vla:2251
            Exit Function
        Case 312
            vla_step_text = "Align cell A6 center. [line 592]" ' vla:2252
            Exit Function
        Case 313
            vla_step_text = "Wrap text in range C1:C5. [line 593]" ' vla:2253
            Exit Function
        Case 314
            vla_step_text = "Unwrap text in range C1:C5. [line 594]" ' vla:2254
            Exit Function
        Case 315
            vla_step_text = "Merge range D1:D3. [line 595]" ' vla:2255
            Exit Function
        Case 316
            vla_step_text = "Unmerge range D1:D3. [line 596]" ' vla:2256
            Exit Function
        Case 317
            vla_step_text = "Set height of row 1 to 30. [line 601]" ' vla:2257
            Exit Function
        Case 318
            vla_step_text = "Set width of column A to 20. [line 602]" ' vla:2258
            Exit Function
        Case 319
            vla_step_text = "Insert column before B. [line 603]" ' vla:2259
            Exit Function
        Case 320
            vla_step_text = "Hide column C. [line 604]" ' vla:2260
            Exit Function
        Case 321
            vla_step_text = "Unhide column C. [line 605]" ' vla:2261
            Exit Function
        Case 322
            vla_step_text = "Delete column B. [line 606]" ' vla:2262
            Exit Function
        Case 323
            vla_step_text = "Insert row at 5. [line 607]" ' vla:2263
            Exit Function
        Case 324
            vla_step_text = "Insert a row. [line 608]" ' vla:2264
            Exit Function
        Case 325
            vla_step_text = "Delete row 2. [line 609]" ' vla:2265
            Exit Function
        Case 326
            vla_step_text = "Delete the third row. [line 610]" ' vla:2266
            Exit Function
        Case 327
            vla_step_text = "Hide row 10. [line 611]" ' vla:2267
            Exit Function
        Case 328
            vla_step_text = "Unhide row 10. [line 612]" ' vla:2268
            Exit Function
        Case 329
            vla_step_text = "Freeze top row. [line 613]" ' vla:2269
            Exit Function
        Case 330
            vla_step_text = "Unfreeze panes. [line 614]" ' vla:2270
            Exit Function
        Case 331
            vla_step_text = "Fit column A. [line 615]" ' vla:2271
            Exit Function
        Case 332
            vla_step_text = "Fit all columns. [line 616]" ' vla:2272
            Exit Function
        Case 333
            vla_step_text = "Put ""Region"" into cell G1. [line 620]" ' vla:2273
            Exit Function
        Case 334
            vla_step_text = "Put ""Amount"" into cell H1. [line 621]" ' vla:2274
            Exit Function
        Case 335
            vla_step_text = "Put ""Notes"" into cell I1. [line 622]" ' vla:2275
            Exit Function
        Case 336
            vla_step_text = "Put ""West"" into cell G2. [line 623]" ' vla:2276
            Exit Function
        Case 337
            vla_step_text = "Put 100 into cell H2. [line 624]" ' vla:2277
            Exit Function
        Case 338
            vla_step_text = "Put ""ok"" into cell I2. [line 625]" ' vla:2278
            Exit Function
        Case 339
            vla_step_text = "Put ""East"" into cell G3. [line 626]" ' vla:2279
            Exit Function
        Case 340
            vla_step_text = "Put 250 into cell H3. [line 627]" ' vla:2280
            Exit Function
        Case 341
            vla_step_text = "Put ""ok"" into cell I3. [line 628]" ' vla:2281
            Exit Function
        Case 342
            vla_step_text = "Put ""West"" into cell G4. [line 629]" ' vla:2282
            Exit Function
        Case 343
            vla_step_text = "Put 100 into cell H4. [line 630]" ' vla:2283
            Exit Function
        Case 344
            vla_step_text = "Put ""dup"" into cell I4. [line 631]" ' vla:2284
            Exit Function
        Case 345
            vla_step_text = "Sort range G1:I4 by column H1 descending. [line 632]" ' vla:2285
            Exit Function
        Case 346
            vla_step_text = "Keep only rows of range G1:I4 where column 1 is ""West"". [line 633]" ' vla:2286
            Exit Function
        Case 347
            vla_step_text = "Show all rows. [line 634]" ' vla:2287
            Exit Function
        Case 348
            vla_step_text = "Replace ""dup"" with ""ok"" in range G1:I4. [line 635]" ' vla:2288
            Exit Function
        Case 349
            vla_step_text = "Remove duplicates from range G1:I4 by column 1. [line 636]" ' vla:2289
            Exit Function
        Case 350
            vla_step_text = "Replace ""ok"" with ""fine"" in column I. [line 637]" ' vla:2290
            Exit Function
        Case 351
            vla_step_text = "Convert range G1:I4 to values. [line 638]" ' vla:2291
            Exit Function
        Case 352
            vla_step_text = "Clear formatting of range G1:I4. [line 639]" ' vla:2292
            Exit Function
        Case 353
            vla_step_text = "Name range G1:I4 as demo_table. [line 640]" ' vla:2293
            Exit Function
        Case 354
            vla_step_text = "Copy range G1:I4 to range K1:M4. [line 641]" ' vla:2294
            Exit Function
        Case 355
            vla_step_text = "Set tab-color of sheet Demo to hot-pink. [line 645]" ' vla:2295
            Exit Function
        Case 356
            vla_step_text = "Protect this sheet with password ""demo123"". [line 646]" ' vla:2296
            Exit Function
        Case 357
            vla_step_text = "Unprotect this sheet with password ""demo123"". [line 647]" ' vla:2297
            Exit Function
        Case 358
            vla_step_text = "Try: [line 675]" ' vla:2298
            Exit Function
        Case 359
            vla_step_text = "Delete sheet GStruct. [line 676]" ' vla:2299
            Exit Function
        Case 360
            vla_step_text = "Work on sheet GStruct. [line 678]" ' vla:2300
            Exit Function
        Case 361
            vla_step_text = "Hide row 3. [line 680]" ' vla:2301
            Exit Function
        Case 362
            vla_step_text = "Hide column B. [line 681]" ' vla:2302
            Exit Function
        Case 363
            vla_step_text = "Unhide all rows and columns. [line 682]" ' vla:2303
            Exit Function
        Case 364
            vla_step_text = "Set font size of cell A6 to 36. [line 684]" ' vla:2304
            Exit Function
        Case 365
            vla_step_text = "Fit row 6. [line 685]" ' vla:2305
            Exit Function
        Case 366
            vla_step_text = "Group rows 10 through 12. [line 687]" ' vla:2306
            Exit Function
        Case 367
            vla_step_text = "Group rows 14 through 16. [line 688]" ' vla:2307
            Exit Function
        Case 368
            vla_step_text = "Ungroup rows 14 through 16. [line 689]" ' vla:2308
            Exit Function
        Case 369
            vla_step_text = "Put ""before-insert"" into cell A20. [line 691]" ' vla:2309
            Exit Function
        Case 370
            vla_step_text = "Insert 3 rows at row 20. [line 692]" ' vla:2310
            Exit Function
        Case 371
            vla_step_text = "Put ""before-delete"" into cell A40. [line 694]" ' vla:2311
            Exit Function
        Case 372
            vla_step_text = "Delete rows 38 through 39. [line 695]" ' vla:2312
            Exit Function
        Case 373
            vla_step_text = "Put ""marker-b"" into cell B50. [line 697]" ' vla:2313
            Exit Function
        Case 374
            vla_step_text = "Put ""marker-c"" into cell C50. [line 698]" ' vla:2314
            Exit Function
        Case 375
            vla_step_text = "Put ""marker-d"" into cell D50. [line 699]" ' vla:2315
            Exit Function
        Case 376
            vla_step_text = "Move column B before column D. [line 700]" ' vla:2316
            Exit Function
        Case 377
            vla_step_text = "Freeze the first 2 rows. [line 702]" ' vla:2317
            Exit Function
        Case 378
            vla_step_text = "Make cell A60 bold. [line 705]" ' vla:2318
            Exit Function
        Case 379
            vla_step_text = "Make cell A60 red. [line 706]" ' vla:2319
            Exit Function
        Case 380
            vla_step_text = "Copy formatting of A60:A60 to C60:C60. [line 707]" ' vla:2320
            Exit Function
        Case 381
            vla_step_text = "Make C62:C62 look like A60:A60. [line 708]" ' vla:2321
            Exit Function
        Case 382
            vla_step_text = "Put formula ""=5*2"" into cell B64. [line 710]" ' vla:2322
            Exit Function
        Case 383
            vla_step_text = "Copy formulas of B64 to D64. [line 711]" ' vla:2323
            Exit Function
        Case 384
            vla_step_text = "Set width of column F to 33. [line 713]" ' vla:2324
            Exit Function
        Case 385
            vla_step_text = "Copy column widths of F1:F1 to H1:H1. [line 714]" ' vla:2325
            Exit Function
        Case 386
            vla_step_text = "Put 1 into cell A68. [line 716]" ' vla:2326
            Exit Function
        Case 387
            vla_step_text = "Put 2 into cell A69. [line 717]" ' vla:2327
            Exit Function
        Case 388
            vla_step_text = "Put 3 into cell A70. [line 718]" ' vla:2328
            Exit Function
        Case 389
            vla_step_text = "Copy range A68:A70 to C68 transposed. [line 719]" ' vla:2329
            Exit Function
        Case 390
            vla_step_text = "Put ""cutme"" into cell A72. [line 721]" ' vla:2330
            Exit Function
        Case 391
            vla_step_text = "Move range A72:A72 to C72. [line 722]" ' vla:2331
            Exit Function
        Case 392
            vla_step_text = "Put ""rowdata"" into cell A74. [line 724]" ' vla:2332
            Exit Function
        Case 393
            vla_step_text = "Copy row 74 to row 76. [line 725]" ' vla:2333
            Exit Function
        Case 394
            vla_step_text = "Put ""clearme"" into cell A80. [line 728]" ' vla:2334
            Exit Function
        Case 395
            vla_step_text = "Make cell A80 bold. [line 729]" ' vla:2335
            Exit Function
        Case 396
            vla_step_text = "Make cell A80 red. [line 730]" ' vla:2336
            Exit Function
        Case 397
            vla_step_text = "Clear everything from A80:B80. [line 731]" ' vla:2337
            Exit Function
        Case 398
            vla_step_text = "Put ""m1"" into cell A84. [line 733]" ' vla:2338
            Exit Function
        Case 399
            vla_step_text = "Put ""m2"" into cell A85. [line 734]" ' vla:2339
            Exit Function
        Case 400
            vla_step_text = "Put ""m3"" into cell A86. [line 735]" ' vla:2340
            Exit Function
        Case 401
            vla_step_text = "Put ""m4"" into cell A87. [line 736]" ' vla:2341
            Exit Function
        Case 402
            vla_step_text = "Put ""m5"" into cell A88. [line 737]" ' vla:2342
            Exit Function
        Case 403
            vla_step_text = "Delete A84:A86 and shift cells up. [line 738]" ' vla:2343
            Exit Function
        Case 404
            vla_step_text = "Put ""x1"" into cell B90. [line 740]" ' vla:2344
            Exit Function
        Case 405
            vla_step_text = "Put ""x2"" into cell C90. [line 741]" ' vla:2345
            Exit Function
        Case 406
            vla_step_text = "Put ""x3"" into cell D90. [line 742]" ' vla:2346
            Exit Function
        Case 407
            vla_step_text = "Put ""rightdata"" into cell E90. [line 743]" ' vla:2347
            Exit Function
        Case 408
            vla_step_text = "Delete B90:D90 and shift cells left. [line 744]" ' vla:2348
            Exit Function
        Case 409
            vla_step_text = "Put 1 into cell A94. [line 746]" ' vla:2349
            Exit Function
        Case 410
            vla_step_text = "Put 2 into cell B94. [line 747]" ' vla:2350
            Exit Function
        Case 411
            vla_step_text = "Put 3 into cell A95. [line 748]" ' vla:2351
            Exit Function
        Case 412
            vla_step_text = "Put 5 into cell A96. [line 749]" ' vla:2352
            Exit Function
        Case 413
            vla_step_text = "Put 6 into cell B96. [line 750]" ' vla:2353
            Exit Function
        Case 414
            vla_step_text = "Delete blank rows in A94:B96. [line 751]" ' vla:2354
            Exit Function
        Case 415
            vla_step_text = "Put 1 into cell A150. [line 754]" ' vla:2355
            Exit Function
        Case 416
            vla_step_text = "Put 2 into cell A151. [line 755]" ' vla:2356
            Exit Function
        Case 417
            vla_step_text = "Put 3 into cell A152. [line 756]" ' vla:2357
            Exit Function
        Case 418
            vla_step_text = "Put 4 into cell A153. [line 757]" ' vla:2358
            Exit Function
        Case 419
            vla_step_text = "Band every other row of A150:D153 ""#D9D9D9"". [line 758]" ' vla:2359
            Exit Function
        Case 420
            vla_step_text = "Make row 165 a header row. [line 759]" ' vla:2360
            Exit Function
        Case 421
            vla_step_text = "Fill A170:A179 with a series starting at 1. [line 762]" ' vla:2361
            Exit Function
        Case 422
            vla_step_text = "Fill A180:A184 with a series starting at 1 with step 2. [line 763]" ' vla:2362
            Exit Function
        Case 423
            vla_step_text = "Fill A190:A194 with a growth series starting at 2. [line 764]" ' vla:2363
            Exit Function
        Case 424
            vla_step_text = "Fill A200:A203 with a growth series starting at 2 with step 3. [line 765]" ' vla:2364
            Exit Function
        Case 425
            vla_step_text = "Make cell Z500 bold. [line 777]" ' vla:2365
            Exit Function
        Case 426
            vla_step_text = "Clear everything from Z500. [line 778]" ' vla:2366
            Exit Function
        Case 427
            vla_step_text = "Remove trailing empty rows and columns. [line 779]" ' vla:2367
            Exit Function
        Case 428
            vla_step_text = "Go to sheet Output. [line 781]" ' vla:2368
            Exit Function
        Case 429
            vla_step_text = "Tidy-up. [line 783]" ' vla:2369
            Exit Function
        Case 430
            vla_step_text = "Turn on screen updating. [line 784]" ' vla:2370
            Exit Function
        Case 431
            vla_step_text = "Log ""report finished"". [line 785]" ' vla:2371
            Exit Function
        Case Else
            vla_step_text = "an unknown step" ' vla:2372
            Exit Function
    End Select
End Function


