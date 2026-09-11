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

Public Sub main()
    Dim counter As Variant ' vla:66
    Dim grand As Variant ' vla:67
    Dim results As Variant ' vla:68
    Dim r As Variant ' vla:69
    Dim biggest As Variant ' vla:70
    Dim probe As Variant ' vla:71
    Dim round_check As Variant ' vla:72
    Dim thousand_check As Variant ' vla:73
    Dim f_last As Variant ' vla:74
    Dim echo_row As Variant ' vla:75
    Dim check_row As Variant ' vla:76
    Dim stripe_row As Variant ' vla:77
    Dim back_row As Variant ' vla:78
    Dim search_row As Variant ' vla:79
    Dim f As Variant ' vla:80
    Dim list_count As Variant ' vla:81
    Dim region As Variant ' vla:82
    Dim fuel As Variant ' vla:83
    Dim fee As Variant ' vla:84
    Dim full_commission As Variant ' vla:85
    Dim default_commission As Variant ' vla:86
    Dim pick_check As Variant ' vla:87
    Dim ax_price As Variant ' vla:88
    Dim key_count As Variant ' vla:89
    Dim price_sum As Variant ' vla:90
    Dim k As Variant ' vla:91
    Dim pair As Variant ' vla:92
    Dim qty_values As Variant ' vla:93
    Dim q As Variant ' vla:94
    On Error GoTo vla_fail ' vla:95
    vla_step = 1 ' vla:96
    If vlatraceon() Then ' vla:97
        Call vlatracestep(1, vla_step_text(1)) ' vla:97
    End If
    Call vlaensuresheet("output") ' vla:99 src:18
    Call worksheets("output").activate
    vla_step = 2 ' vla:100
    If vlatraceon() Then ' vla:101
        Call vlatracestep(2, vla_step_text(2)) ' vla:101
    End If
    application.screenupdating = False
    vla_step = 3 ' vla:104
    If vlatraceon() Then ' vla:105
        Call vlatracestep(3, vla_step_text(3)) ' vla:105
    End If
    Dim total As Double ' vla:107 src:21
    vla_step = 4 ' vla:108
    If vlatraceon() Then ' vla:109
        Call vlatracestep(4, vla_step_text(4)) ' vla:109
    End If
    total = 0 ' vla:111 src:22
    vla_step = 5 ' vla:112
    If vlatraceon() Then ' vla:113
        Call vlatracestep(5, vla_step_text(5)) ' vla:113
    End If
    range("a1") = "Test Report" ' vla:115 src:23
    vla_step = 6 ' vla:116
    If vlatraceon() Then ' vla:117
        Call vlatracestep(6, vla_step_text(6)) ' vla:117
    End If
    range("a1").font.bold = True
    vla_step = 7 ' vla:120
    If vlatraceon() Then ' vla:121
        Call vlatracestep(7, vla_step_text(7)) ' vla:121
    End If
    range("a1").font.size = 14
    vla_step = 8 ' vla:124
    If vlatraceon() Then ' vla:125
        Call vlatracestep(8, vla_step_text(8)) ' vla:125
    End If
    range("d1") = date() ' vla:127 src:26
    vla_step = 9 ' vla:128
    If vlatraceon() Then ' vla:129
        Call vlatracestep(9, vla_step_text(9)) ' vla:129
    End If
    For counter = 1 To 5
        vla_step = 10 ' vla:132 src:28
        If vlatraceon() Then ' vla:133 src:28
            Call vlatracestep(10, vla_step_text(10)) ' vla:133 src:28
        End If
        total = (total + counter)
        vla_step = 11 ' vla:136 src:28
        If vlatraceon() Then ' vla:137 src:28
            Call vlatracestep(11, vla_step_text(11)) ' vla:137 src:28
        End If
        If ((counter Mod 2) = 0) Then ' vla:139 src:30
            Debug.Print ("even step " & counter) ' vla:139 src:30
        End If
    Next counter
    vla_step = 12 ' vla:140
    If vlatraceon() Then ' vla:141
        Call vlatracestep(12, vla_step_text(12)) ' vla:141
    End If
    range("b2") = total ' vla:143 src:32
    vla_step = 13 ' vla:144
    If vlatraceon() Then ' vla:145
        Call vlatracestep(13, vla_step_text(13)) ' vla:145
    End If
    range("b3").Formula2 = "=B2*2"
    vla_step = 14 ' vla:148
    If vlatraceon() Then ' vla:149
        Call vlatracestep(14, vla_step_text(14)) ' vla:149
    End If
    range("b4") = application.worksheetfunction.sum(range("B2:B3")) ' vla:151 src:37
    vla_step = 15 ' vla:152
    If vlatraceon() Then ' vla:153
        Call vlatracestep(15, vla_step_text(15)) ' vla:153
    End If
    grand = application.worksheetfunction.sum(range("b2:b3")) ' vla:155 src:38
    vla_step = 16 ' vla:156
    If vlatraceon() Then ' vla:157
        Call vlatracestep(16, vla_step_text(16)) ' vla:157
    End If
    Debug.Print ("grand is " & grand) ' vla:159 src:39
    vla_step = 17 ' vla:160
    If vlatraceon() Then ' vla:161
        Call vlatracestep(17, vla_step_text(17)) ' vla:161
    End If
    If (grand > 40) Then ' vla:163 src:43
        vla_step = 18 ' vla:165 src:43
        If vlatraceon() Then ' vla:166 src:43
            Call vlatracestep(18, vla_step_text(18)) ' vla:166 src:43
        End If
        range("c4") = "PASS" ' vla:168 src:44
        vla_step = 19 ' vla:169 src:43
        If vlatraceon() Then ' vla:170 src:43
            Call vlatracestep(19, vla_step_text(19)) ' vla:170 src:43
        End If
        For counter = 1 To 2
            vla_step = 20 ' vla:173 src:45
            If vlatraceon() Then ' vla:174 src:45
                Call vlatracestep(20, vla_step_text(20)) ' vla:174 src:45
            End If
            Debug.Print ("pass check " & counter) ' vla:176 src:46
        Next counter
        vla_step = 21 ' vla:177 src:43
        If vlatraceon() Then ' vla:178 src:43
            Call vlatracestep(21, vla_step_text(21)) ' vla:178 src:43
        End If
        range("c4").font.bold = True
    Else
        vla_step = 22 ' vla:182 src:43
        If vlatraceon() Then ' vla:183 src:43
            Call vlatracestep(22, vla_step_text(22)) ' vla:183 src:43
        End If
        range("c4") = "CHECK" ' vla:185 src:51
    End If
    vla_step = 23 ' vla:186
    If vlatraceon() Then ' vla:187
        Call vlatracestep(23, vla_step_text(23)) ' vla:187
    End If
    If (grand >= 45) Then ' vla:189 src:53
        range("c4").interior.color = vbyellow
    End If
    vla_step = 24 ' vla:190
    If vlatraceon() Then ' vla:191
        Call vlatracestep(24, vla_step_text(24)) ' vla:191
    End If
    Dim label As String ' vla:193 src:54
    vla_step = 25 ' vla:194
    If vlatraceon() Then ' vla:195
        Call vlatracestep(25, vla_step_text(25)) ' vla:195
    End If
    label = ("Total: " & total) ' vla:197 src:55
    vla_step = 26 ' vla:198
    If vlatraceon() Then ' vla:199
        Call vlatracestep(26, vla_step_text(26)) ' vla:199
    End If
    range("a6") = label ' vla:201 src:56
    vla_step = 27 ' vla:202
    If vlatraceon() Then ' vla:203
        Call vlatracestep(27, vla_step_text(27)) ' vla:203
    End If
    range("a6").font.color = vlacolor(hot_pink)
    vla_step = 28 ' vla:206
    If vlatraceon() Then ' vla:207
        Call vlatracestep(28, vla_step_text(28)) ' vla:207
    End If
    Set results = range("b2:b4") ' vla:209 src:58
    vla_step = 29 ' vla:210
    If vlatraceon() Then ' vla:211
        Call vlatracestep(29, vla_step_text(29)) ' vla:211
    End If
    For Each r In results ' vla:213 src:59
        Debug.Print r ' vla:213 src:59
    Next r
    vla_step = 30 ' vla:214
    If vlatraceon() Then ' vla:215
        Call vlatracestep(30, vla_step_text(30)) ' vla:215
    End If
    biggest = application.worksheetfunction.max(results) ' vla:217 src:60
    vla_step = 31 ' vla:218
    If vlatraceon() Then ' vla:219
        Call vlatracestep(31, vla_step_text(31)) ' vla:219
    End If
    Debug.Print ((("largest result is " & biggest) & ", label length ") & len(label)) ' vla:221 src:61
    vla_step = 32 ' vla:222
    If vlatraceon() Then ' vla:223
        Call vlatracestep(32, vla_step_text(32)) ' vla:223
    End If
    For counter = 1 To 3
        vla_step = 33 ' vla:226 src:64
        If vlatraceon() Then ' vla:227 src:64
            Call vlatracestep(33, vla_step_text(33)) ' vla:227 src:64
        End If
        cells(counter, "e") = (counter * 10) ' vla:229 src:65
    Next counter
    vla_step = 34 ' vla:230
    If vlatraceon() Then ' vla:231
        Call vlatracestep(34, vla_step_text(34)) ' vla:231
    End If
    probe = cells(2, "e") ' vla:233 src:67
    vla_step = 35 ' vla:234
    If vlatraceon() Then ' vla:235
        Call vlatracestep(35, vla_step_text(35)) ' vla:235
    End If
    Dim num_col_check As String ' vla:237 src:68
    vla_step = 36 ' vla:238
    If vlatraceon() Then ' vla:239
        Call vlatracestep(36, vla_step_text(36)) ' vla:239
    End If
    If (cells(2, 5) = 20) Then ' vla:241 src:69
        num_col_check = "yes" ' vla:241 src:69
    End If
    vla_step = 37 ' vla:242
    If vlatraceon() Then ' vla:243
        Call vlatracestep(37, vla_step_text(37)) ' vla:243
    End If
    Dim value_word_check As String ' vla:245 src:70
    vla_step = 38 ' vla:246
    If vlatraceon() Then ' vla:247
        Call vlatracestep(38, vla_step_text(38)) ' vla:247
    End If
    If (cells(3, 5) = 30) Then ' vla:249 src:71
        value_word_check = "ok" ' vla:249 src:71
    End If
    vla_step = 39 ' vla:250
    If vlatraceon() Then ' vla:251
        Call vlatracestep(39, vla_step_text(39)) ' vla:251
    End If
    range("output!h16") = "bang" ' vla:253 src:72
    vla_step = 40 ' vla:254
    If vlatraceon() Then ' vla:255
        Call vlatracestep(40, vla_step_text(40)) ' vla:255
    End If
    range("a1").horizontalalignment = xlcenter
    vla_step = 41 ' vla:258
    If vlatraceon() Then ' vla:259
        Call vlatracestep(41, vla_step_text(41)) ' vla:259
    End If
    columns("d").columnwidth = 24
    vla_step = 42 ' vla:262
    If vlatraceon() Then ' vla:263
        Call vlatracestep(42, vla_step_text(42)) ' vla:263
    End If
    round_check = round(3.14159, 2) ' vla:265 src:75
    vla_step = 43 ' vla:266
    If vlatraceon() Then ' vla:267
        Call vlatracestep(43, vla_step_text(43)) ' vla:267
    End If
    Debug.Print ("rounded is " & round_check) ' vla:269 src:76
    vla_step = 44 ' vla:270
    If vlatraceon() Then ' vla:271
        Call vlatracestep(44, vla_step_text(44)) ' vla:271
    End If
    thousand_check = (1000 + 500) ' vla:273 src:77
    vla_step = 45 ' vla:274
    If vlatraceon() Then ' vla:275
        Call vlatracestep(45, vla_step_text(45)) ' vla:275
    End If
    Debug.Print ("thousands read as " & thousand_check) ' vla:277 src:78
    vla_step = 46 ' vla:278
    If vlatraceon() Then ' vla:279
        Call vlatracestep(46, vla_step_text(46)) ' vla:279
    End If
    Debug.Print ("probe is " & probe) ' vla:281 src:79
    vla_step = 47 ' vla:282
    If vlatraceon() Then ' vla:283
        Call vlatracestep(47, vla_step_text(47)) ' vla:283
    End If
    Dim countdown As Double ' vla:285 src:81
    vla_step = 48 ' vla:286
    If vlatraceon() Then ' vla:287
        Call vlatracestep(48, vla_step_text(48)) ' vla:287
    End If
    countdown = 3 ' vla:289 src:82
    vla_step = 49 ' vla:290
    If vlatraceon() Then ' vla:291
        Call vlatracestep(49, vla_step_text(49)) ' vla:291
    End If
    Do While (countdown > 0) ' vla:293 src:83
        vla_step = 50 ' vla:294 src:83
        If vlatraceon() Then ' vla:295 src:83
            Call vlatracestep(50, vla_step_text(50)) ' vla:295 src:83
        End If
        Debug.Print ("countdown " & countdown) ' vla:297 src:84
        vla_step = 51 ' vla:298 src:83
        If vlatraceon() Then ' vla:299 src:83
            Call vlatracestep(51, vla_step_text(51)) ' vla:299 src:83
        End If
        countdown = (countdown - 1) ' vla:301 src:85
    Loop
    vla_step = 53 ' vla:302
    If vlatraceon() Then ' vla:303
        Call vlatracestep(53, vla_step_text(53)) ' vla:303
    End If
    Call stamp ' vla:305 src:92
    vla_step = 54 ' vla:306
    If vlatraceon() Then ' vla:307
        Call vlatracestep(54, vla_step_text(54)) ' vla:307
    End If
    Call stamp(row_number:=2, value:="beta") ' vla:309 src:93
    vla_step = 55 ' vla:310
    If vlatraceon() Then ' vla:311
        Call vlatracestep(55, vla_step_text(55)) ' vla:311
    End If
    f_last = cells(rows.count, "f").end(xlup).row ' vla:313 src:94
    vla_step = 56 ' vla:314
    If vlatraceon() Then ' vla:315
        Call vlatracestep(56, vla_step_text(56)) ' vla:315
    End If
    Debug.Print ("column F filled to row " & f_last) ' vla:317 src:95
    vla_step = 57 ' vla:318
    If vlatraceon() Then ' vla:319
        Call vlatracestep(57, vla_step_text(57)) ' vla:319
    End If
    For echo_row = 1 To f_last ' vla:321 src:96
        Debug.Print ("echo " & echo_row) ' vla:321 src:96
    Next echo_row
    vla_step = 58 ' vla:322
    If vlatraceon() Then ' vla:323
        Call vlatracestep(58, vla_step_text(58)) ' vla:323
    End If
    For check_row = 1 To f_last ' vla:325 src:97
        vla_step = 59 ' vla:326 src:97
        If vlatraceon() Then ' vla:327 src:97
            Call vlatracestep(59, vla_step_text(59)) ' vla:327 src:97
        End If
        If (instr(1, cells(check_row, "f"), "ok", vbtextcompare) > 0) Then ' vla:329 src:98
            cells(check_row, "f").font.bold = True
        End If
    Next check_row
    vla_step = 60 ' vla:330
    If vlatraceon() Then ' vla:331
        Call vlatracestep(60, vla_step_text(60)) ' vla:331
    End If
    For stripe_row = 1 To f_last Step 2 ' vla:333 src:103
        vla_step = 61 ' vla:334 src:103
        If vlatraceon() Then ' vla:335 src:103
            Call vlatracestep(61, vla_step_text(61)) ' vla:335 src:103
        End If
        cells(stripe_row, "f").font.bold = True
    Next stripe_row
    vla_step = 62 ' vla:338
    If vlatraceon() Then ' vla:339
        Call vlatracestep(62, vla_step_text(62)) ' vla:339
    End If
    For back_row = f_last To 1 Step (0 - 2) ' vla:341 src:106
        vla_step = 63 ' vla:342 src:106
        If vlatraceon() Then ' vla:343 src:106
            Call vlatracestep(63, vla_step_text(63)) ' vla:343 src:106
        End If
        Debug.Print ("back-row " & back_row) ' vla:345 src:107
    Next back_row
    vla_step = 64 ' vla:346
    If vlatraceon() Then ' vla:347
        Call vlatracestep(64, vla_step_text(64)) ' vla:347
    End If
    For search_row = 1 To 10 ' vla:349 src:111
        vla_step = 65 ' vla:350 src:111
        If vlatraceon() Then ' vla:351 src:111
            Call vlatracestep(65, vla_step_text(65)) ' vla:351 src:111
        End If
        If (search_row = 3) Then ' vla:353 src:112
            Exit For ' vla:353 src:112
        End If
    Next search_row
    vla_step = 66 ' vla:354
    If vlatraceon() Then ' vla:355
        Call vlatracestep(66, vla_step_text(66)) ' vla:355
    End If
    Debug.Print ("stopped at " & search_row) ' vla:357 src:114
    vla_step = 67 ' vla:358
    If vlatraceon() Then ' vla:359
        Call vlatracestep(67, vla_step_text(67)) ' vla:359
    End If
    Dim found_items As Collection ' vla:361 src:119
    Set found_items = New Collection ' vla:361 src:119
    vla_step = 68 ' vla:362
    If vlatraceon() Then ' vla:363
        Call vlatracestep(68, vla_step_text(68)) ' vla:363
    End If
    For counter = 1 To 4
        vla_step = 69 ' vla:366 src:120
        If vlatraceon() Then ' vla:367 src:120
            Call vlatracestep(69, vla_step_text(69)) ' vla:367 src:120
        End If
        If (counter > 2) Then ' vla:369 src:121
            Call found_items.add((counter * 100)) ' vla:369 src:121
        End If
    Next counter
    vla_step = 70 ' vla:370
    If vlatraceon() Then ' vla:371
        Call vlatracestep(70, vla_step_text(70)) ' vla:371
    End If
    For Each f In found_items ' vla:373 src:123
        Debug.Print ("found " & f) ' vla:373 src:123
    Next f
    vla_step = 71 ' vla:374
    If vlatraceon() Then ' vla:375
        Call vlatracestep(71, vla_step_text(71)) ' vla:375
    End If
    list_count = vlacount(found_items) ' vla:377 src:124
    vla_step = 72 ' vla:378
    If vlatraceon() Then ' vla:379
        Call vlatracestep(72, vla_step_text(72)) ' vla:379
    End If
    Debug.Print ("list holds " & list_count) ' vla:381 src:125
    vla_step = 73 ' vla:382
    If vlatraceon() Then ' vla:383
        Call vlatracestep(73, vla_step_text(73)) ' vla:383
    End If
    Dim verdict As String ' vla:385 src:131
    vla_step = 74 ' vla:386
    If vlatraceon() Then ' vla:387
        Call vlatracestep(74, vla_step_text(74)) ' vla:387
    End If
    If (grand > 100) Then ' vla:389 src:132
        vla_step = 75 ' vla:391 src:132
        If vlatraceon() Then ' vla:392 src:132
            Call vlatracestep(75, vla_step_text(75)) ' vla:392 src:132
        End If
        verdict = "huge" ' vla:394 src:133
    ElseIf (grand > 40) Then
        vla_step = 76 ' vla:396 src:132
        If vlatraceon() Then ' vla:397 src:132
            Call vlatracestep(76, vla_step_text(76)) ' vla:397 src:132
        End If
        verdict = "solid" ' vla:399 src:136
    Else
        vla_step = 77 ' vla:401 src:132
        If vlatraceon() Then ' vla:402 src:132
            Call vlatracestep(77, vla_step_text(77)) ' vla:402 src:132
        End If
        verdict = "small" ' vla:404 src:139
    End If
    vla_step = 78 ' vla:405
    If vlatraceon() Then ' vla:406
        Call vlatracestep(78, vla_step_text(78)) ' vla:406
    End If
    Dim region_label As String ' vla:408 src:141
    vla_step = 79 ' vla:409
    If vlatraceon() Then ' vla:410
        Call vlatracestep(79, vla_step_text(79)) ' vla:410
    End If
    region = "South" ' vla:412 src:142
    vla_step = 80 ' vla:413
    If vlatraceon() Then ' vla:414
        Call vlatracestep(80, vla_step_text(80)) ' vla:414
    End If
    Select Case region ' vla:416 src:143
        Case "North"
            vla_step = 81 ' vla:418 src:143
            If vlatraceon() Then ' vla:419 src:143
                Call vlatracestep(81, vla_step_text(81)) ' vla:419 src:143
            End If
            region_label = "cold" ' vla:421 src:144
        Case "South", "East"
            vla_step = 82 ' vla:423 src:143
            If vlatraceon() Then ' vla:424 src:143
                Call vlatracestep(82, vla_step_text(82)) ' vla:424 src:143
            End If
            region_label = "warm" ' vla:426 src:147
        Case Else
            vla_step = 83 ' vla:428 src:143
            If vlatraceon() Then ' vla:429 src:143
                Call vlatracestep(83, vla_step_text(83)) ' vla:429 src:143
            End If
            region_label = "unknown" ' vla:431 src:150
    End Select
    vla_step = 84 ' vla:432
    If vlatraceon() Then ' vla:433
        Call vlatracestep(84, vla_step_text(84)) ' vla:433
    End If
    Debug.Print ((("verdict " & verdict) & ", region ") & region_label) ' vla:435 src:152
    vla_step = 85 ' vla:436
    If vlatraceon() Then ' vla:437
        Call vlatracestep(85, vla_step_text(85)) ' vla:437
    End If
    Dim until_count As Double ' vla:439 src:154
    vla_step = 86 ' vla:440
    If vlatraceon() Then ' vla:441
        Call vlatracestep(86, vla_step_text(86)) ' vla:441
    End If
    fuel = 3 ' vla:443 src:155
    vla_step = 87 ' vla:444
    If vlatraceon() Then ' vla:445
        Call vlatracestep(87, vla_step_text(87)) ' vla:445
    End If
    Do Until (fuel = 0) ' vla:447 src:156
        vla_step = 88 ' vla:448 src:156
        If vlatraceon() Then ' vla:449 src:156
            Call vlatracestep(88, vla_step_text(88)) ' vla:449 src:156
        End If
        fuel = (fuel - 1) ' vla:451 src:157
        vla_step = 89 ' vla:452 src:156
        If vlatraceon() Then ' vla:453 src:156
            Call vlatracestep(89, vla_step_text(89)) ' vla:453 src:156
        End If
        until_count = (until_count + 1)
    Loop
    vla_step = 90 ' vla:456
    If vlatraceon() Then ' vla:457
        Call vlatracestep(90, vla_step_text(90)) ' vla:457
    End If
    Debug.Print (("repeat-until ran " & until_count) & " times") ' vla:459 src:160
    vla_step = 91 ' vla:460
    If vlatraceon() Then ' vla:461
        Call vlatracestep(91, vla_step_text(91)) ' vla:461
    End If
    Dim rescue As String ' vla:463 src:166
    vla_step = 92 ' vla:464
    If vlatraceon() Then ' vla:465
        Call vlatracestep(92, vla_step_text(92)) ' vla:465
    End If
    On Error GoTo vla_tryf_1 ' vla:467 src:167
    vla_step = 93 ' vla:468 src:167
    If vlatraceon() Then ' vla:469 src:167
        Call vlatracestep(93, vla_step_text(93)) ' vla:469 src:167
    End If
    Call worksheets("nowhere-land").activate
    vla_step = 94 ' vla:472 src:167
    If vlatraceon() Then ' vla:473 src:167
        Call vlatracestep(94, vla_step_text(94)) ' vla:473 src:167
    End If
    rescue = "unreachable" ' vla:475 src:169
    GoTo vla_tryd_1 ' vla:476 src:167
vla_tryf_1: ' vla:477 src:167
    vla_problem = err.description ' vla:478 src:167
    Resume vla_tryr_1 ' vla:479 src:167
vla_tryr_1: ' vla:480 src:167
    On Error GoTo vla_fail ' vla:481 src:167
    vla_step = 95 ' vla:482 src:167
    If vlatraceon() Then ' vla:483 src:167
        Call vlatracestep(95, vla_step_text(95)) ' vla:483 src:167
    End If
    Debug.Print ("the problem was " & vla_problem) ' vla:485 src:172
    vla_step = 96 ' vla:486 src:167
    If vlatraceon() Then ' vla:487 src:167
        Call vlatracestep(96, vla_step_text(96)) ' vla:487 src:167
    End If
    If (Not (len(trim((vla_problem & ""))) = 0)) Then ' vla:489 src:173
        rescue = "rescued" ' vla:489 src:173
    End If
vla_tryd_1: ' vla:490 src:167
    On Error GoTo vla_fail ' vla:491 src:167
    vla_step = 97 ' vla:492
    If vlatraceon() Then ' vla:493
        Call vlatracestep(97, vla_step_text(97)) ' vla:493
    End If
    Dim risk_free As Double ' vla:495 src:175
    vla_step = 98 ' vla:496
    If vlatraceon() Then ' vla:497
        Call vlatracestep(98, vla_step_text(98)) ' vla:497
    End If
    On Error GoTo vla_tryf_2 ' vla:499 src:176
    vla_step = 99 ' vla:500 src:176
    If vlatraceon() Then ' vla:501 src:176
        Call vlatracestep(99, vla_step_text(99)) ' vla:501 src:176
    End If
    risk_free = 7 ' vla:503 src:177
    GoTo vla_tryd_2 ' vla:504 src:176
vla_tryf_2: ' vla:505 src:176
    vla_problem = err.description ' vla:506 src:176
    Resume vla_tryr_2 ' vla:507 src:176
vla_tryr_2: ' vla:508 src:176
    On Error GoTo vla_fail ' vla:509 src:176
    vla_step = 100 ' vla:510 src:176
    If vlatraceon() Then ' vla:511 src:176
        Call vlatracestep(100, vla_step_text(100)) ' vla:511 src:176
    End If
    risk_free = -1 ' vla:513 src:180
vla_tryd_2: ' vla:514 src:176
    On Error GoTo vla_fail ' vla:515 src:176
    vla_step = 101 ' vla:516
    If vlatraceon() Then ' vla:517
        Call vlatracestep(101, vla_step_text(101)) ' vla:517
    End If
    Debug.Print ((("rescue " & rescue) & ", risk-free ") & risk_free) ' vla:519 src:182
    vla_step = 103 ' vla:520
    If vlatraceon() Then ' vla:521
        Call vlatracestep(103, vla_step_text(103)) ' vla:521
    End If
    fee = tax(100) ' vla:523 src:189
    vla_step = 104 ' vla:524
    If vlatraceon() Then ' vla:525
        Call vlatracestep(104, vla_step_text(104)) ' vla:525
    End If
    Debug.Print ("fee is " & fee) ' vla:527 src:190
    vla_step = 105 ' vla:528
    If vlatraceon() Then ' vla:529
        Call vlatracestep(105, vla_step_text(105)) ' vla:529
    End If
    Dim fee_size As String ' vla:531 src:191
    vla_step = 106 ' vla:532
    If vlatraceon() Then ' vla:533
        Call vlatracestep(106, vla_step_text(106)) ' vla:533
    End If
    If (tax(50) > 3) Then ' vla:535 src:192
        vla_step = 107 ' vla:537 src:192
        If vlatraceon() Then ' vla:538 src:192
            Call vlatracestep(107, vla_step_text(107)) ' vla:538 src:192
        End If
        fee_size = "big" ' vla:540 src:193
    Else
        vla_step = 108 ' vla:542 src:192
        If vlatraceon() Then ' vla:543 src:192
            Call vlatracestep(108, vla_step_text(108)) ' vla:543 src:192
        End If
        fee_size = "small" ' vla:545 src:196
    End If
    vla_step = 113 ' vla:546
    If vlatraceon() Then ' vla:547
        Call vlatracestep(113, vla_step_text(113)) ' vla:547
    End If
    full_commission = commission(sale:=2000, rate:=(10 / 100)) ' vla:549 src:208
    vla_step = 114 ' vla:550
    If vlatraceon() Then ' vla:551
        Call vlatracestep(114, vla_step_text(114)) ' vla:551
    End If
    default_commission = commission(sale:=600) ' vla:553 src:209
    vla_step = 115 ' vla:554
    If vlatraceon() Then ' vla:555
        Call vlatracestep(115, vla_step_text(115)) ' vla:555
    End If
    Debug.Print ((("commissions " & full_commission) & " / ") & default_commission) ' vla:557 src:210
    vla_step = 116 ' vla:558
    If vlatraceon() Then ' vla:559
        Call vlatracestep(116, vla_step_text(116)) ' vla:559
    End If
    On Error GoTo vla_tryf_3 ' vla:561 src:219
    vla_step = 117 ' vla:562 src:219
    If vlatraceon() Then ' vla:563 src:219
        Call vlatracestep(117, vla_step_text(117)) ' vla:563 src:219
    End If
    Call vlachecksheetname("Q1 Data")
    Call vlachecksheetabsent("Q1 Data")
    Call worksheets.add
    activesheet.name = "Q1 Data"
    GoTo vla_tryd_3 ' vla:566 src:219
vla_tryf_3: ' vla:567 src:219
    vla_problem = err.description ' vla:568 src:219
    Resume vla_tryr_3 ' vla:569 src:219
vla_tryr_3: ' vla:570 src:219
    On Error GoTo vla_fail ' vla:571 src:219
vla_tryd_3: ' vla:572 src:219
    On Error GoTo vla_fail ' vla:573 src:219
    vla_step = 118 ' vla:574
    If vlatraceon() Then ' vla:575
        Call vlatracestep(118, vla_step_text(118)) ' vla:575
    End If
    Call worksheets("output").activate
    vla_step = 119 ' vla:578
    If vlatraceon() Then ' vla:579
        Call vlatracestep(119, vla_step_text(119)) ' vla:579
    End If
    range("'Q1 Data'!A1") = "spaced" ' vla:581 src:223
    vla_step = 120 ' vla:582
    If vlatraceon() Then ' vla:583
        Call vlatracestep(120, vla_step_text(120)) ' vla:583
    End If
    Dim spaced_check As String ' vla:585 src:224
    vla_step = 121 ' vla:586
    If vlatraceon() Then ' vla:587
        Call vlatracestep(121, vla_step_text(121)) ' vla:587
    End If
    spaced_check = range("'Q1 Data'!A1") ' vla:589 src:225
    vla_step = 123 ' vla:590
    If vlatraceon() Then ' vla:591
        Call vlatracestep(123, vla_step_text(123)) ' vla:591
    End If
    Dim growth_check As Double ' vla:593 src:237
    vla_step = 124 ' vla:594
    If vlatraceon() Then ' vla:595
        Call vlatracestep(124, vla_step_text(124)) ' vla:595
    End If
    growth_check = 200 ' vla:597 src:238
    vla_step = 125 ' vla:598
    If vlatraceon() Then ' vla:599
        Call vlatracestep(125, vla_step_text(125)) ' vla:599
    End If
    growth_check = (growth_check * (1 + (10 / 100))) ' vla:601 src:239
    vla_step = 126 ' vla:602
    If vlatraceon() Then ' vla:603
        Call vlatracestep(126, vla_step_text(126)) ' vla:603
    End If
    growth_check = (growth_check * (1 + (50 / 100))) ' vla:605 src:240
    vla_step = 127 ' vla:606
    If vlatraceon() Then ' vla:607
        Call vlatracestep(127, vla_step_text(127)) ' vla:607
    End If
    growth_check = (growth_check * (1 + (100 / 100))) ' vla:609 src:241
    vla_step = 128 ' vla:610
    If vlatraceon() Then ' vla:611
        Call vlatracestep(128, vla_step_text(128)) ' vla:611
    End If
    growth_check = (growth_check * (1 - (75 / 100))) ' vla:613 src:242
    vla_step = 129 ' vla:614
    If vlatraceon() Then ' vla:615
        Call vlatracestep(129, vla_step_text(129)) ' vla:615
    End If
    pick_check = vlaitem(found_items, 2) ' vla:617 src:243
    vla_step = 130 ' vla:618
    If vlatraceon() Then ' vla:619
        Call vlatracestep(130, vla_step_text(130)) ' vla:619
    End If
    Debug.Print ((((("grew to " & growth_check) & ", picked ") & pick_check) & ", first ") & vlafirst(found_items)) ' vla:621 src:244
    vla_step = 131 ' vla:622
    If vlatraceon() Then ' vla:623
        Call vlatracestep(131, vla_step_text(131)) ' vla:623
    End If
    Dim quote_check As String ' vla:625 src:245
    vla_step = 132 ' vla:626
    If vlatraceon() Then ' vla:627
        Call vlatracestep(132, vla_step_text(132)) ' vla:627
    End If
    quote_check = "He said ""ok""" ' vla:629 src:246
    vla_step = 133 ' vla:630
    If vlatraceon() Then ' vla:631
        Call vlatracestep(133, vla_step_text(133)) ' vla:631
    End If
    Dim prices As Object ' vla:633 src:260
    Set prices = vladictnew() ' vla:633 src:260
    vla_step = 134 ' vla:634
    If vlatraceon() Then ' vla:635
        Call vlatracestep(134, vla_step_text(134)) ' vla:635
    End If
    Call vladictset(prices, "ax-7", 100) ' vla:637 src:261
    vla_step = 135 ' vla:638
    If vlatraceon() Then ' vla:639
        Call vlatracestep(135, vla_step_text(135)) ' vla:639
    End If
    Call vladictset(prices, "bx-2", 250) ' vla:641 src:262
    vla_step = 136 ' vla:642
    If vlatraceon() Then ' vla:643
        Call vlatracestep(136, vla_step_text(136)) ' vla:643
    End If
    Call vladictset(prices, "AX-7", 120) ' vla:645 src:263
    vla_step = 137 ' vla:646
    If vlatraceon() Then ' vla:647
        Call vlatracestep(137, vla_step_text(137)) ' vla:647
    End If
    ax_price = vladictget(prices, "ax-7") ' vla:649 src:264
    vla_step = 138 ' vla:650
    If vlatraceon() Then ' vla:651
        Call vlatracestep(138, vla_step_text(138)) ' vla:651
    End If
    key_count = vlacount(vladictkeys(prices)) ' vla:653 src:265
    vla_step = 139 ' vla:654
    If vlatraceon() Then ' vla:655
        Call vlatracestep(139, vla_step_text(139)) ' vla:655
    End If
    price_sum = (vladictget(prices, "ax-7") + vladictget(prices, "bx-2")) ' vla:657 src:266
    vla_step = 140 ' vla:658
    If vlatraceon() Then ' vla:659
        Call vlatracestep(140, vla_step_text(140)) ' vla:659
    End If
    Dim key_list As String ' vla:661 src:267
    vla_step = 141 ' vla:662
    If vlatraceon() Then ' vla:663
        Call vlatracestep(141, vla_step_text(141)) ' vla:663
    End If
    For Each k In vladictkeys(prices) ' vla:665 src:268
        vla_step = 142 ' vla:666 src:268
        If vlatraceon() Then ' vla:667 src:268
            Call vlatracestep(142, vla_step_text(142)) ' vla:667 src:268
        End If
        key_list = (key_list & k) ' vla:669 src:269
    Next k
    vla_step = 143 ' vla:670
    If vlatraceon() Then ' vla:671
        Call vlatracestep(143, vla_step_text(143)) ' vla:671
    End If
    Dim price_verdict As String ' vla:673 src:271
    vla_step = 144 ' vla:674
    If vlatraceon() Then ' vla:675
        Call vlatracestep(144, vla_step_text(144)) ' vla:675
    End If
    If (vladictget(prices, "bx-2") > 200) Then ' vla:677 src:272
        price_verdict = "steep" ' vla:677 src:272
    End If
    vla_step = 145 ' vla:678
    If vlatraceon() Then ' vla:679
        Call vlatracestep(145, vla_step_text(145)) ' vla:679
    End If
    Dim pair_trace As String ' vla:681 src:273
    vla_step = 146 ' vla:682
    If vlatraceon() Then ' vla:683
        Call vlatracestep(146, vla_step_text(146)) ' vla:683
    End If
    For Each pair In vladictpairs(prices) ' vla:685 src:274
        vla_step = 147 ' vla:686 src:274
        If vlatraceon() Then ' vla:687 src:274
            Call vlatracestep(147, vla_step_text(147)) ' vla:687 src:274
        End If
        If (vlapairvalue(pair) > 200) Then ' vla:689 src:275
            pair_trace = (pair_trace & vlapairkey(pair)) ' vla:689 src:275
        End If
    Next pair
    vla_step = 148 ' vla:690
    If vlatraceon() Then ' vla:691
        Call vlatracestep(148, vla_step_text(148)) ' vla:691
    End If
    Debug.Print ((((("lookup: ax " & ax_price) & ", keys ") & key_list) & ", pairs ") & pair_trace) ' vla:693 src:277
    vla_step = 149 ' vla:694
    If vlatraceon() Then ' vla:695
        Call vlatracestep(149, vla_step_text(149)) ' vla:695
    End If
    range("j1") = "Item" ' vla:697 src:284
    vla_step = 150 ' vla:698
    If vlatraceon() Then ' vla:699
        Call vlatracestep(150, vla_step_text(150)) ' vla:699
    End If
    range("k1") = "Amount" ' vla:701 src:285
    vla_step = 151 ' vla:702
    If vlatraceon() Then ' vla:703
        Call vlatracestep(151, vla_step_text(151)) ' vla:703
    End If
    range("j2") = "Widget" ' vla:705 src:286
    vla_step = 152 ' vla:706
    If vlatraceon() Then ' vla:707
        Call vlatracestep(152, vla_step_text(152)) ' vla:707
    End If
    range("k2") = 10 ' vla:709 src:287
    vla_step = 153 ' vla:710
    If vlatraceon() Then ' vla:711
        Call vlatracestep(153, vla_step_text(153)) ' vla:711
    End If
    range("j3") = "Gadget" ' vla:713 src:288
    vla_step = 154 ' vla:714
    If vlatraceon() Then ' vla:715
        Call vlatracestep(154, vla_step_text(154)) ' vla:715
    End If
    range("k3") = 20 ' vla:717 src:289
    vla_step = 155 ' vla:718
    If vlatraceon() Then ' vla:719
        Call vlatracestep(155, vla_step_text(155)) ' vla:719
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j1:k3"), xllistobjecthasheaders:=xlyes).name = "salestable"
    vla_step = 156 ' vla:722
    If vlatraceon() Then ' vla:723
        Call vlatracestep(156, vla_step_text(156)) ' vla:723
    End If
    activesheet.listobjects("salestable").tablestyle = "tablestylemedium9"
    vla_step = 157 ' vla:726
    If vlatraceon() Then ' vla:727
        Call vlatracestep(157, vla_step_text(157)) ' vla:727
    End If
    activesheet.listobjects("salestable").showtotals = True
    vla_step = 158 ' vla:730
    If vlatraceon() Then ' vla:731
        Call vlatracestep(158, vla_step_text(158)) ' vla:731
    End If
    range("j5") = "X" ' vla:733 src:296
    vla_step = 159 ' vla:734
    If vlatraceon() Then ' vla:735
        Call vlatracestep(159, vla_step_text(159)) ' vla:735
    End If
    range("j6") = "Y" ' vla:737 src:297
    vla_step = 160 ' vla:738
    If vlatraceon() Then ' vla:739
        Call vlatracestep(160, vla_step_text(160)) ' vla:739
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j5:j6"), xllistobjecthasheaders:=xlyes).name = "quiettable"
    vla_step = 161 ' vla:742
    If vlatraceon() Then ' vla:743
        Call vlatracestep(161, vla_step_text(161)) ' vla:743
    End If
    activesheet.listobjects("quiettable").showtotals = True
    vla_step = 162 ' vla:746
    If vlatraceon() Then ' vla:747
        Call vlatracestep(162, vla_step_text(162)) ' vla:747
    End If
    activesheet.listobjects("quiettable").showtotals = False
    vla_step = 163 ' vla:750
    If vlatraceon() Then ' vla:751
        Call vlatracestep(163, vla_step_text(163)) ' vla:751
    End If
    range("j8") = "A" ' vla:753 src:304
    vla_step = 164 ' vla:754
    If vlatraceon() Then ' vla:755
        Call vlatracestep(164, vla_step_text(164)) ' vla:755
    End If
    range("j9") = "B" ' vla:757 src:305
    vla_step = 165 ' vla:758
    If vlatraceon() Then ' vla:759
        Call vlatracestep(165, vla_step_text(165)) ' vla:759
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j8:j9"), xllistobjecthasheaders:=xlyes).name = "temptable"
    vla_step = 166 ' vla:762
    If vlatraceon() Then ' vla:763
        Call vlatracestep(166, vla_step_text(166)) ' vla:763
    End If
    Call activesheet.listobjects("temptable").unlist
    vla_step = 167 ' vla:766
    If vlatraceon() Then ' vla:767
        Call vlatracestep(167, vla_step_text(167)) ' vla:767
    End If
    range("j11") = "Item" ' vla:769 src:312
    vla_step = 168 ' vla:770
    If vlatraceon() Then ' vla:771
        Call vlatracestep(168, vla_step_text(168)) ' vla:771
    End If
    range("k11") = "Qty" ' vla:773 src:313
    vla_step = 169 ' vla:774
    If vlatraceon() Then ' vla:775
        Call vlatracestep(169, vla_step_text(169)) ' vla:775
    End If
    range("j12") = "Bolt" ' vla:777 src:314
    vla_step = 170 ' vla:778
    If vlatraceon() Then ' vla:779
        Call vlatracestep(170, vla_step_text(170)) ' vla:779
    End If
    range("k12") = 5 ' vla:781 src:315
    vla_step = 171 ' vla:782
    If vlatraceon() Then ' vla:783
        Call vlatracestep(171, vla_step_text(171)) ' vla:783
    End If
    range("j13") = "Nut" ' vla:785 src:316
    vla_step = 172 ' vla:786
    If vlatraceon() Then ' vla:787
        Call vlatracestep(172, vla_step_text(172)) ' vla:787
    End If
    range("k13") = 8 ' vla:789 src:317
    vla_step = 173 ' vla:790
    If vlatraceon() Then ' vla:791
        Call vlatracestep(173, vla_step_text(173)) ' vla:791
    End If
    activesheet.listobjects.add(sourcetype:=xlsrcrange, source:=range("j11:k13"), xllistobjecthasheaders:=xlyes).name = "edittable"
    vla_step = 174 ' vla:794
    If vlatraceon() Then ' vla:795
        Call vlatracestep(174, vla_step_text(174)) ' vla:795
    End If
    Call activesheet.listobjects("edittable").listrows.add
    vla_step = 175 ' vla:798
    If vlatraceon() Then ' vla:799
        Call vlatracestep(175, vla_step_text(175)) ' vla:799
    End If
    Call activesheet.listobjects("edittable").listrows(1).delete
    vla_step = 176 ' vla:802
    If vlatraceon() Then ' vla:803
        Call vlatracestep(176, vla_step_text(176)) ' vla:803
    End If
    Set qty_values = activesheet.listobjects("edittable").listcolumns("qty").databodyrange
    vla_step = 177 ' vla:806
    If vlatraceon() Then ' vla:807
        Call vlatracestep(177, vla_step_text(177)) ' vla:807
    End If
    For Each q In qty_values ' vla:809 src:322
        Debug.Print q ' vla:809 src:322
    Next q
    vla_step = 178 ' vla:810
    If vlatraceon() Then ' vla:811
        Call vlatracestep(178, vla_step_text(178)) ' vla:811
    End If
    range("n1") = "Region" ' vla:813 src:332
    vla_step = 179 ' vla:814
    If vlatraceon() Then ' vla:815
        Call vlatracestep(179, vla_step_text(179)) ' vla:815
    End If
    range("o1") = "Product" ' vla:817 src:333
    vla_step = 180 ' vla:818
    If vlatraceon() Then ' vla:819
        Call vlatracestep(180, vla_step_text(180)) ' vla:819
    End If
    range("p1") = "Segment" ' vla:821 src:334
    vla_step = 181 ' vla:822
    If vlatraceon() Then ' vla:823
        Call vlatracestep(181, vla_step_text(181)) ' vla:823
    End If
    range("q1") = "Channel" ' vla:825 src:335
    vla_step = 182 ' vla:826
    If vlatraceon() Then ' vla:827
        Call vlatracestep(182, vla_step_text(182)) ' vla:827
    End If
    range("r1") = "Units" ' vla:829 src:336
    vla_step = 183 ' vla:830
    If vlatraceon() Then ' vla:831
        Call vlatracestep(183, vla_step_text(183)) ' vla:831
    End If
    range("s1") = "Revenue" ' vla:833 src:337
    vla_step = 184 ' vla:834
    If vlatraceon() Then ' vla:835
        Call vlatracestep(184, vla_step_text(184)) ' vla:835
    End If
    range("n2") = "North" ' vla:837 src:338
    vla_step = 185 ' vla:838
    If vlatraceon() Then ' vla:839
        Call vlatracestep(185, vla_step_text(185)) ' vla:839
    End If
    range("o2") = "Widget" ' vla:841 src:339
    vla_step = 186 ' vla:842
    If vlatraceon() Then ' vla:843
        Call vlatracestep(186, vla_step_text(186)) ' vla:843
    End If
    range("p2") = "Retail" ' vla:845 src:340
    vla_step = 187 ' vla:846
    If vlatraceon() Then ' vla:847
        Call vlatracestep(187, vla_step_text(187)) ' vla:847
    End If
    range("q2") = "Online" ' vla:849 src:341
    vla_step = 188 ' vla:850
    If vlatraceon() Then ' vla:851
        Call vlatracestep(188, vla_step_text(188)) ' vla:851
    End If
    range("r2") = 10 ' vla:853 src:342
    vla_step = 189 ' vla:854
    If vlatraceon() Then ' vla:855
        Call vlatracestep(189, vla_step_text(189)) ' vla:855
    End If
    range("s2") = 500 ' vla:857 src:343
    vla_step = 190 ' vla:858
    If vlatraceon() Then ' vla:859
        Call vlatracestep(190, vla_step_text(190)) ' vla:859
    End If
    range("n3") = "North" ' vla:861 src:344
    vla_step = 191 ' vla:862
    If vlatraceon() Then ' vla:863
        Call vlatracestep(191, vla_step_text(191)) ' vla:863
    End If
    range("o3") = "Gadget" ' vla:865 src:345
    vla_step = 192 ' vla:866
    If vlatraceon() Then ' vla:867
        Call vlatracestep(192, vla_step_text(192)) ' vla:867
    End If
    range("p3") = "Wholesale" ' vla:869 src:346
    vla_step = 193 ' vla:870
    If vlatraceon() Then ' vla:871
        Call vlatracestep(193, vla_step_text(193)) ' vla:871
    End If
    range("q3") = "Store" ' vla:873 src:347
    vla_step = 194 ' vla:874
    If vlatraceon() Then ' vla:875
        Call vlatracestep(194, vla_step_text(194)) ' vla:875
    End If
    range("r3") = 5 ' vla:877 src:348
    vla_step = 195 ' vla:878
    If vlatraceon() Then ' vla:879
        Call vlatracestep(195, vla_step_text(195)) ' vla:879
    End If
    range("s3") = 200 ' vla:881 src:349
    vla_step = 196 ' vla:882
    If vlatraceon() Then ' vla:883
        Call vlatracestep(196, vla_step_text(196)) ' vla:883
    End If
    range("n4") = "South" ' vla:885 src:350
    vla_step = 197 ' vla:886
    If vlatraceon() Then ' vla:887
        Call vlatracestep(197, vla_step_text(197)) ' vla:887
    End If
    range("o4") = "Widget" ' vla:889 src:351
    vla_step = 198 ' vla:890
    If vlatraceon() Then ' vla:891
        Call vlatracestep(198, vla_step_text(198)) ' vla:891
    End If
    range("p4") = "Wholesale" ' vla:893 src:352
    vla_step = 199 ' vla:894
    If vlatraceon() Then ' vla:895
        Call vlatracestep(199, vla_step_text(199)) ' vla:895
    End If
    range("q4") = "Online" ' vla:897 src:353
    vla_step = 200 ' vla:898
    If vlatraceon() Then ' vla:899
        Call vlatracestep(200, vla_step_text(200)) ' vla:899
    End If
    range("r4") = 20 ' vla:901 src:354
    vla_step = 201 ' vla:902
    If vlatraceon() Then ' vla:903
        Call vlatracestep(201, vla_step_text(201)) ' vla:903
    End If
    range("s4") = 900 ' vla:905 src:355
    vla_step = 202 ' vla:906
    If vlatraceon() Then ' vla:907
        Call vlatracestep(202, vla_step_text(202)) ' vla:907
    End If
    range("n5") = "South" ' vla:909 src:356
    vla_step = 203 ' vla:910
    If vlatraceon() Then ' vla:911
        Call vlatracestep(203, vla_step_text(203)) ' vla:911
    End If
    range("o5") = "Gadget" ' vla:913 src:357
    vla_step = 204 ' vla:914
    If vlatraceon() Then ' vla:915
        Call vlatracestep(204, vla_step_text(204)) ' vla:915
    End If
    range("p5") = "Retail" ' vla:917 src:358
    vla_step = 205 ' vla:918
    If vlatraceon() Then ' vla:919
        Call vlatracestep(205, vla_step_text(205)) ' vla:919
    End If
    range("q5") = "Store" ' vla:921 src:359
    vla_step = 206 ' vla:922
    If vlatraceon() Then ' vla:923
        Call vlatracestep(206, vla_step_text(206)) ' vla:923
    End If
    range("r5") = 8 ' vla:925 src:360
    vla_step = 207 ' vla:926
    If vlatraceon() Then ' vla:927
        Call vlatracestep(207, vla_step_text(207)) ' vla:927
    End If
    range("s5") = 300 ' vla:929 src:361
    vla_step = 208 ' vla:930
    If vlatraceon() Then ' vla:931
        Call vlatracestep(208, vla_step_text(208)) ' vla:931
    End If
    range("n6") = "East" ' vla:933 src:362
    vla_step = 209 ' vla:934
    If vlatraceon() Then ' vla:935
        Call vlatracestep(209, vla_step_text(209)) ' vla:935
    End If
    range("o6") = "Widget" ' vla:937 src:363
    vla_step = 210 ' vla:938
    If vlatraceon() Then ' vla:939
        Call vlatracestep(210, vla_step_text(210)) ' vla:939
    End If
    range("p6") = "Retail" ' vla:941 src:364
    vla_step = 211 ' vla:942
    If vlatraceon() Then ' vla:943
        Call vlatracestep(211, vla_step_text(211)) ' vla:943
    End If
    range("q6") = "Online" ' vla:945 src:365
    vla_step = 212 ' vla:946
    If vlatraceon() Then ' vla:947
        Call vlatracestep(212, vla_step_text(212)) ' vla:947
    End If
    range("r6") = 12 ' vla:949 src:366
    vla_step = 213 ' vla:950
    If vlatraceon() Then ' vla:951
        Call vlatracestep(213, vla_step_text(213)) ' vla:951
    End If
    range("s6") = 600 ' vla:953 src:367
    vla_step = 214 ' vla:954
    If vlatraceon() Then ' vla:955
        Call vlatracestep(214, vla_step_text(214)) ' vla:955
    End If
    range("n7") = "East" ' vla:957 src:368
    vla_step = 215 ' vla:958
    If vlatraceon() Then ' vla:959
        Call vlatracestep(215, vla_step_text(215)) ' vla:959
    End If
    range("o7") = "Gadget" ' vla:961 src:369
    vla_step = 216 ' vla:962
    If vlatraceon() Then ' vla:963
        Call vlatracestep(216, vla_step_text(216)) ' vla:963
    End If
    range("p7") = "Wholesale" ' vla:965 src:370
    vla_step = 217 ' vla:966
    If vlatraceon() Then ' vla:967
        Call vlatracestep(217, vla_step_text(217)) ' vla:967
    End If
    range("q7") = "Store" ' vla:969 src:371
    vla_step = 218 ' vla:970
    If vlatraceon() Then ' vla:971
        Call vlatracestep(218, vla_step_text(218)) ' vla:971
    End If
    range("r7") = 6 ' vla:973 src:372
    vla_step = 219 ' vla:974
    If vlatraceon() Then ' vla:975
        Call vlatracestep(219, vla_step_text(219)) ' vla:975
    End If
    range("s7") = 250 ' vla:977 src:373
    vla_step = 220 ' vla:978
    If vlatraceon() Then ' vla:979
        Call vlatracestep(220, vla_step_text(220)) ' vla:979
    End If
    Call vlapivotcreate(range("n1:s7"), range("u1"), "salespivot")
    vla_step = 221 ' vla:982
    If vlatraceon() Then ' vla:983
        Call vlatracestep(221, vla_step_text(221)) ' vla:983
    End If
    Call vlapivotsetorientation("salespivot", Array("region", "product"), "row")
    vla_step = 222 ' vla:986
    If vlatraceon() Then ' vla:987
        Call vlatracestep(222, vla_step_text(222)) ' vla:987
    End If
    Call vlapivotsetorientation("salespivot", Array("segment"), "column")
    vla_step = 223 ' vla:990
    If vlatraceon() Then ' vla:991
        Call vlatracestep(223, vla_step_text(223)) ' vla:991
    End If
    Call vlapivotsetorientation("salespivot", Array("channel"), "filter")
    vla_step = 224 ' vla:994
    If vlatraceon() Then ' vla:995
        Call vlatracestep(224, vla_step_text(224)) ' vla:995
    End If
    Call vlapivotaddvalues("salespivot", Array("revenue"), "sum")
    vla_step = 225 ' vla:998
    If vlatraceon() Then ' vla:999
        Call vlatracestep(225, vla_step_text(225)) ' vla:999
    End If
    Call vlapivotaddvalues("salespivot", Array("revenue", "units"), "count")
    vla_step = 226 ' vla:1002
    If vlatraceon() Then ' vla:1003
        Call vlatracestep(226, vla_step_text(226)) ' vla:1003
    End If
    Call vlapivotaddvalues("salespivot", Array("units"), "average")
    vla_step = 227 ' vla:1006
    If vlatraceon() Then ' vla:1007
        Call vlatracestep(227, vla_step_text(227)) ' vla:1007
    End If
    Call vlapivotcreate(range("n1:s7"), range("n20"), "fullpivot")
    Call vlapivotsetorientation("fullpivot", Array("region", "product", "channel"), "row")
    Call vlapivotsetorientation("fullpivot", Array("segment"), "column")
    Call vlapivotaddvalues("fullpivot", Array("revenue", "units"), "sum")
    vla_step = 228 ' vla:1010
    If vlatraceon() Then ' vla:1011
        Call vlatracestep(228, vla_step_text(228)) ' vla:1011
    End If
    Call vlapivotrefresh("salespivot")
    vla_step = 229 ' vla:1014
    If vlatraceon() Then ' vla:1015
        Call vlatracestep(229, vla_step_text(229)) ' vla:1015
    End If
    Call vlapivotrefreshall
    vla_step = 230 ' vla:1018
    If vlatraceon() Then ' vla:1019
        Call vlatracestep(230, vla_step_text(230)) ' vla:1019
    End If
    Call vlapivotsetshowdetail("salespivot", Array("region"), False)
    vla_step = 231 ' vla:1022
    If vlatraceon() Then ' vla:1023
        Call vlatracestep(231, vla_step_text(231)) ' vla:1023
    End If
    Call vlapivotsetshowdetail("fullpivot", Array("product"), False)
    vla_step = 232 ' vla:1026
    If vlatraceon() Then ' vla:1027
        Call vlatracestep(232, vla_step_text(232)) ' vla:1027
    End If
    Call vlapivotsetshowdetail("fullpivot", Array("product"), True)
    vla_step = 233 ' vla:1030
    If vlatraceon() Then ' vla:1031
        Call vlatracestep(233, vla_step_text(233)) ' vla:1031
    End If
    Call vlapivotsetrowlayout("salespivot", "tabular")
    vla_step = 234 ' vla:1034
    If vlatraceon() Then ' vla:1035
        Call vlatracestep(234, vla_step_text(234)) ' vla:1035
    End If
    Call vlapivotsetrowlayout("salespivot", "compact")
    vla_step = 235 ' vla:1038
    If vlatraceon() Then ' vla:1039
        Call vlatracestep(235, vla_step_text(235)) ' vla:1039
    End If
    Call vlapivotsetrowlayout("fullpivot", "outline")
    vla_step = 236 ' vla:1042
    If vlatraceon() Then ' vla:1043
        Call vlatracestep(236, vla_step_text(236)) ' vla:1043
    End If
    Call vlapivotsetsubtotals("salespivot", Array("region"), False)
    vla_step = 237 ' vla:1046
    If vlatraceon() Then ' vla:1047
        Call vlatracestep(237, vla_step_text(237)) ' vla:1047
    End If
    Call vlapivotsetsubtotals("fullpivot", Array("product"), False)
    vla_step = 238 ' vla:1050
    If vlatraceon() Then ' vla:1051
        Call vlatracestep(238, vla_step_text(238)) ' vla:1051
    End If
    Call vlapivotsetsubtotals("fullpivot", Array("product"), True)
    vla_step = 239 ' vla:1054
    If vlatraceon() Then ' vla:1055
        Call vlatracestep(239, vla_step_text(239)) ' vla:1055
    End If
    Call vlapivotsetblankline("salespivot", Array("region"), True)
    vla_step = 240 ' vla:1058
    If vlatraceon() Then ' vla:1059
        Call vlatracestep(240, vla_step_text(240)) ' vla:1059
    End If
    Call vlapivotsetblankline("fullpivot", Array("product"), True)
    vla_step = 241 ' vla:1062
    If vlatraceon() Then ' vla:1063
        Call vlatracestep(241, vla_step_text(241)) ' vla:1063
    End If
    Call vlapivotsetblankline("fullpivot", Array("product"), False)
    vla_step = 242 ' vla:1066
    If vlatraceon() Then ' vla:1067
        Call vlatracestep(242, vla_step_text(242)) ' vla:1067
    End If
    Call vlapivotsort("salespivot", "region", "descending", "")
    vla_step = 243 ' vla:1070
    If vlatraceon() Then ' vla:1071
        Call vlatracestep(243, vla_step_text(243)) ' vla:1071
    End If
    Call vlapivotsort("salespivot", "region", "ascending", "")
    vla_step = 244 ' vla:1074
    If vlatraceon() Then ' vla:1075
        Call vlatracestep(244, vla_step_text(244)) ' vla:1075
    End If
    Call vlapivotsort("fullpivot", "product", "descending", "revenue")
    vla_step = 245 ' vla:1078
    If vlatraceon() Then ' vla:1079
        Call vlatracestep(245, vla_step_text(245)) ' vla:1079
    End If
    Call vlapivotsort("fullpivot", "product", "ascending", "revenue")
    vla_step = 246 ' vla:1082
    If vlatraceon() Then ' vla:1083
        Call vlatracestep(246, vla_step_text(246)) ' vla:1083
    End If
    Call vlapivotcreate(range("n1:s7"), range("n200"), "renamemepivot")
    vla_step = 247 ' vla:1086
    If vlatraceon() Then ' vla:1087
        Call vlatracestep(247, vla_step_text(247)) ' vla:1087
    End If
    Call vlapivotrename("renamemepivot", "renamedpivot")
    vla_step = 248 ' vla:1090
    If vlatraceon() Then ' vla:1091
        Call vlatracestep(248, vla_step_text(248)) ' vla:1091
    End If
    Call vlapivotcreate(range("n1:s7"), range("n220"), "clearmepivot")
    vla_step = 249 ' vla:1094
    If vlatraceon() Then ' vla:1095
        Call vlatracestep(249, vla_step_text(249)) ' vla:1095
    End If
    Call vlapivotsetorientation("clearmepivot", Array("region"), "row")
    vla_step = 250 ' vla:1098
    If vlatraceon() Then ' vla:1099
        Call vlatracestep(250, vla_step_text(250)) ' vla:1099
    End If
    Call vlapivotaddvalues("clearmepivot", Array("revenue"), "sum")
    vla_step = 251 ' vla:1102
    If vlatraceon() Then ' vla:1103
        Call vlatracestep(251, vla_step_text(251)) ' vla:1103
    End If
    Call vlapivotclear("clearmepivot")
    vla_step = 252 ' vla:1106
    If vlatraceon() Then ' vla:1107
        Call vlatracestep(252, vla_step_text(252)) ' vla:1107
    End If
    Call vlapivotcreate(range("n1:s7"), range("n240"), "removefieldmepivot")
    vla_step = 253 ' vla:1110
    If vlatraceon() Then ' vla:1111
        Call vlatracestep(253, vla_step_text(253)) ' vla:1111
    End If
    Call vlapivotsetorientation("removefieldmepivot", Array("region", "product"), "row")
    vla_step = 254 ' vla:1114
    If vlatraceon() Then ' vla:1115
        Call vlatracestep(254, vla_step_text(254)) ' vla:1115
    End If
    Call vlapivotsetorientation("removefieldmepivot", Array("product"), "hidden")
    vla_step = 255 ' vla:1118
    If vlatraceon() Then ' vla:1119
        Call vlatracestep(255, vla_step_text(255)) ' vla:1119
    End If
    range("n8") = "West" ' vla:1121 src:506
    vla_step = 256 ' vla:1122
    If vlatraceon() Then ' vla:1123
        Call vlatracestep(256, vla_step_text(256)) ' vla:1123
    End If
    range("o8") = "Widget" ' vla:1125 src:507
    vla_step = 257 ' vla:1126
    If vlatraceon() Then ' vla:1127
        Call vlatracestep(257, vla_step_text(257)) ' vla:1127
    End If
    range("p8") = "Retail" ' vla:1129 src:508
    vla_step = 258 ' vla:1130
    If vlatraceon() Then ' vla:1131
        Call vlatracestep(258, vla_step_text(258)) ' vla:1131
    End If
    range("q8") = "Online" ' vla:1133 src:509
    vla_step = 259 ' vla:1134
    If vlatraceon() Then ' vla:1135
        Call vlatracestep(259, vla_step_text(259)) ' vla:1135
    End If
    range("r8") = 15 ' vla:1137 src:510
    vla_step = 260 ' vla:1138
    If vlatraceon() Then ' vla:1139
        Call vlatracestep(260, vla_step_text(260)) ' vla:1139
    End If
    range("s8") = 700 ' vla:1141 src:511
    vla_step = 261 ' vla:1142
    If vlatraceon() Then ' vla:1143
        Call vlatracestep(261, vla_step_text(261)) ' vla:1143
    End If
    Call vlapivotcreate(range("n1:s7"), range("n260"), "sourcetestpivot")
    vla_step = 262 ' vla:1146
    If vlatraceon() Then ' vla:1147
        Call vlatracestep(262, vla_step_text(262)) ' vla:1147
    End If
    Call vlapivotsetorientation("sourcetestpivot", Array("region"), "row")
    vla_step = 263 ' vla:1150
    If vlatraceon() Then ' vla:1151
        Call vlatracestep(263, vla_step_text(263)) ' vla:1151
    End If
    Call vlapivotchangesource("sourcetestpivot", range("n1:s8"))
    vla_step = 264 ' vla:1154
    If vlatraceon() Then ' vla:1155
        Call vlatracestep(264, vla_step_text(264)) ' vla:1155
    End If
    Call vlapivotcreate(range("n1:s7"), range("n40"), "temppivot")
    vla_step = 265 ' vla:1158
    If vlatraceon() Then ' vla:1159
        Call vlatracestep(265, vla_step_text(265)) ' vla:1159
    End If
    Call vlapivotdelete("temppivot")
    vla_step = 266 ' vla:1162
    If vlatraceon() Then ' vla:1163
        Call vlatracestep(266, vla_step_text(266)) ' vla:1163
    End If
    range("h1") = grand ' vla:1165 src:524
    vla_step = 267 ' vla:1166
    If vlatraceon() Then ' vla:1167
        Call vlatracestep(267, vla_step_text(267)) ' vla:1167
    End If
    range("h2") = biggest ' vla:1169 src:525
    vla_step = 268 ' vla:1170
    If vlatraceon() Then ' vla:1171
        Call vlatracestep(268, vla_step_text(268)) ' vla:1171
    End If
    range("h3") = round_check ' vla:1173 src:526
    vla_step = 269 ' vla:1174
    If vlatraceon() Then ' vla:1175
        Call vlatracestep(269, vla_step_text(269)) ' vla:1175
    End If
    range("h4") = thousand_check ' vla:1177 src:527
    vla_step = 270 ' vla:1178
    If vlatraceon() Then ' vla:1179
        Call vlatracestep(270, vla_step_text(270)) ' vla:1179
    End If
    range("h5") = search_row ' vla:1181 src:528
    vla_step = 271 ' vla:1182
    If vlatraceon() Then ' vla:1183
        Call vlatracestep(271, vla_step_text(271)) ' vla:1183
    End If
    range("h6") = f_last ' vla:1185 src:529
    vla_step = 272 ' vla:1186
    If vlatraceon() Then ' vla:1187
        Call vlatracestep(272, vla_step_text(272)) ' vla:1187
    End If
    range("h7") = list_count ' vla:1189 src:530
    vla_step = 273 ' vla:1190
    If vlatraceon() Then ' vla:1191
        Call vlatracestep(273, vla_step_text(273)) ' vla:1191
    End If
    range("h8") = verdict ' vla:1193 src:531
    vla_step = 274 ' vla:1194
    If vlatraceon() Then ' vla:1195
        Call vlatracestep(274, vla_step_text(274)) ' vla:1195
    End If
    range("h9") = region_label ' vla:1197 src:532
    vla_step = 275 ' vla:1198
    If vlatraceon() Then ' vla:1199
        Call vlatracestep(275, vla_step_text(275)) ' vla:1199
    End If
    range("h10") = until_count ' vla:1201 src:533
    vla_step = 276 ' vla:1202
    If vlatraceon() Then ' vla:1203
        Call vlatracestep(276, vla_step_text(276)) ' vla:1203
    End If
    range("h11") = rescue ' vla:1205 src:534
    vla_step = 277 ' vla:1206
    If vlatraceon() Then ' vla:1207
        Call vlatracestep(277, vla_step_text(277)) ' vla:1207
    End If
    range("h12") = risk_free ' vla:1209 src:535
    vla_step = 278 ' vla:1210
    If vlatraceon() Then ' vla:1211
        Call vlatracestep(278, vla_step_text(278)) ' vla:1211
    End If
    range("h13") = fee ' vla:1213 src:536
    vla_step = 279 ' vla:1214
    If vlatraceon() Then ' vla:1215
        Call vlatracestep(279, vla_step_text(279)) ' vla:1215
    End If
    range("h14") = fee_size ' vla:1217 src:537
    vla_step = 280 ' vla:1218
    If vlatraceon() Then ' vla:1219
        Call vlatracestep(280, vla_step_text(280)) ' vla:1219
    End If
    range("h15") = num_col_check ' vla:1221 src:538
    vla_step = 281 ' vla:1222
    If vlatraceon() Then ' vla:1223
        Call vlatracestep(281, vla_step_text(281)) ' vla:1223
    End If
    range("h17") = value_word_check ' vla:1225 src:539
    vla_step = 282 ' vla:1226
    If vlatraceon() Then ' vla:1227
        Call vlatracestep(282, vla_step_text(282)) ' vla:1227
    End If
    range("h18") = spaced_check ' vla:1229 src:540
    vla_step = 283 ' vla:1230
    If vlatraceon() Then ' vla:1231
        Call vlatracestep(283, vla_step_text(283)) ' vla:1231
    End If
    range("h19") = full_commission ' vla:1233 src:541
    vla_step = 284 ' vla:1234
    If vlatraceon() Then ' vla:1235
        Call vlatracestep(284, vla_step_text(284)) ' vla:1235
    End If
    range("h20") = default_commission ' vla:1237 src:542
    vla_step = 285 ' vla:1238
    If vlatraceon() Then ' vla:1239
        Call vlatracestep(285, vla_step_text(285)) ' vla:1239
    End If
    range("h21") = growth_check ' vla:1241 src:543
    vla_step = 286 ' vla:1242
    If vlatraceon() Then ' vla:1243
        Call vlatracestep(286, vla_step_text(286)) ' vla:1243
    End If
    range("h22") = pick_check ' vla:1245 src:544
    vla_step = 287 ' vla:1246
    If vlatraceon() Then ' vla:1247
        Call vlatracestep(287, vla_step_text(287)) ' vla:1247
    End If
    range("h23") = quote_check ' vla:1249 src:545
    vla_step = 288 ' vla:1250
    If vlatraceon() Then ' vla:1251
        Call vlatracestep(288, vla_step_text(288)) ' vla:1251
    End If
    range("h24") = vat_rate() ' vla:1253 src:546
    vla_step = 289 ' vla:1254
    If vlatraceon() Then ' vla:1255
        Call vlatracestep(289, vla_step_text(289)) ' vla:1255
    End If
    range("h26") = ax_price ' vla:1257 src:547
    vla_step = 290 ' vla:1258
    If vlatraceon() Then ' vla:1259
        Call vlatracestep(290, vla_step_text(290)) ' vla:1259
    End If
    range("h27") = key_count ' vla:1261 src:548
    vla_step = 291 ' vla:1262
    If vlatraceon() Then ' vla:1263
        Call vlatracestep(291, vla_step_text(291)) ' vla:1263
    End If
    range("h28") = key_list ' vla:1265 src:549
    vla_step = 292 ' vla:1266
    If vlatraceon() Then ' vla:1267
        Call vlatracestep(292, vla_step_text(292)) ' vla:1267
    End If
    range("h29") = price_sum ' vla:1269 src:550
    vla_step = 293 ' vla:1270
    If vlatraceon() Then ' vla:1271
        Call vlatracestep(293, vla_step_text(293)) ' vla:1271
    End If
    range("h30") = price_verdict ' vla:1273 src:551
    vla_step = 294 ' vla:1274
    If vlatraceon() Then ' vla:1275
        Call vlatracestep(294, vla_step_text(294)) ' vla:1275
    End If
    range("h31") = pair_trace ' vla:1277 src:552
    vla_step = 295 ' vla:1278
    If vlatraceon() Then ' vla:1279
        Call vlatracestep(295, vla_step_text(295)) ' vla:1279
    End If
    range("h25") = "vla-row" ' vla:1281 src:559
    vla_step = 296 ' vla:1282
    If vlatraceon() Then ' vla:1283
        Call vlatracestep(296, vla_step_text(296)) ' vla:1283
    End If
    For counter = 1 To 2
        vla_step = 297 ' vla:1286 src:560
        If vlatraceon() Then ' vla:1287 src:560
            Call vlatracestep(297, vla_step_text(297)) ' vla:1287 src:560
        End If
        Debug.Print counter ' vla:1289 src:561
    Next counter
    vla_step = 298 ' vla:1290
    If vlatraceon() Then ' vla:1291
        Call vlatracestep(298, vla_step_text(298)) ' vla:1291
    End If
    Call vlaensuresheet("demo") ' vla:1293 src:576
    Call worksheets("demo").activate
    vla_step = 299 ' vla:1294
    If vlatraceon() Then ' vla:1295
        Call vlatracestep(299, vla_step_text(299)) ' vla:1295
    End If
    range("a1").font.italic = True
    vla_step = 300 ' vla:1298
    If vlatraceon() Then ' vla:1299
        Call vlatracestep(300, vla_step_text(300)) ' vla:1299
    End If
    range("a2").interior.color = vbred
    vla_step = 301 ' vla:1302
    If vlatraceon() Then ' vla:1303
        Call vlatracestep(301, vla_step_text(301)) ' vla:1303
    End If
    range("a3").interior.color = vbyellow
    vla_step = 302 ' vla:1306
    If vlatraceon() Then ' vla:1307
        Call vlatracestep(302, vla_step_text(302)) ' vla:1307
    End If
    range("a4").font.color = vlacolor(hot_pink)
    vla_step = 303 ' vla:1310
    If vlatraceon() Then ' vla:1311
        Call vlatracestep(303, vla_step_text(303)) ' vla:1311
    End If
    range("a5").interior.color = vlacolor("#FF69B4")
    vla_step = 304 ' vla:1314
    If vlatraceon() Then ' vla:1315
        Call vlatracestep(304, vla_step_text(304)) ' vla:1315
    End If
    range("a5").interior.colorindex = xlnone
    vla_step = 305 ' vla:1318
    If vlatraceon() Then ' vla:1319
        Call vlatracestep(305, vla_step_text(305)) ' vla:1319
    End If
    range("a1:e10").borders.linestyle = xlcontinuous
    vla_step = 306 ' vla:1322
    If vlatraceon() Then ' vla:1323
        Call vlatracestep(306, vla_step_text(306)) ' vla:1323
    End If
    range("b1").numberformat = "$#,##0.00"
    vla_step = 307 ' vla:1326
    If vlatraceon() Then ' vla:1327
        Call vlatracestep(307, vla_step_text(307)) ' vla:1327
    End If
    range("b2").numberformat = "0.0%"
    vla_step = 308 ' vla:1330
    If vlatraceon() Then ' vla:1331
        Call vlatracestep(308, vla_step_text(308)) ' vla:1331
    End If
    range("b3").numberformat = "mm/dd/yyyy"
    vla_step = 309 ' vla:1334
    If vlatraceon() Then ' vla:1335
        Call vlatracestep(309, vla_step_text(309)) ' vla:1335
    End If
    range("a1").horizontalalignment = xlcenter
    vla_step = 310 ' vla:1338
    If vlatraceon() Then ' vla:1339
        Call vlatracestep(310, vla_step_text(310)) ' vla:1339
    End If
    range("a1:e1").horizontalalignment = xlcenter
    vla_step = 311 ' vla:1342
    If vlatraceon() Then ' vla:1343
        Call vlatracestep(311, vla_step_text(311)) ' vla:1343
    End If
    range("a2:a5").horizontalalignment = xlleft
    vla_step = 312 ' vla:1346
    If vlatraceon() Then ' vla:1347
        Call vlatracestep(312, vla_step_text(312)) ' vla:1347
    End If
    range("a6").horizontalalignment = xlcenter
    vla_step = 313 ' vla:1350
    If vlatraceon() Then ' vla:1351
        Call vlatracestep(313, vla_step_text(313)) ' vla:1351
    End If
    range("c1:c5").wraptext = True
    vla_step = 314 ' vla:1354
    If vlatraceon() Then ' vla:1355
        Call vlatracestep(314, vla_step_text(314)) ' vla:1355
    End If
    range("c1:c5").wraptext = False
    vla_step = 315 ' vla:1358
    If vlatraceon() Then ' vla:1359
        Call vlatracestep(315, vla_step_text(315)) ' vla:1359
    End If
    Call range("d1:d3").merge
    vla_step = 316 ' vla:1362
    If vlatraceon() Then ' vla:1363
        Call vlatracestep(316, vla_step_text(316)) ' vla:1363
    End If
    Call range("d1:d3").unmerge
    vla_step = 317 ' vla:1366
    If vlatraceon() Then ' vla:1367
        Call vlatracestep(317, vla_step_text(317)) ' vla:1367
    End If
    rows(1).rowheight = 30
    vla_step = 318 ' vla:1370
    If vlatraceon() Then ' vla:1371
        Call vlatracestep(318, vla_step_text(318)) ' vla:1371
    End If
    columns("a").columnwidth = 20
    vla_step = 319 ' vla:1374
    If vlatraceon() Then ' vla:1375
        Call vlatracestep(319, vla_step_text(319)) ' vla:1375
    End If
    Call columns("b").insert
    vla_step = 320 ' vla:1378
    If vlatraceon() Then ' vla:1379
        Call vlatracestep(320, vla_step_text(320)) ' vla:1379
    End If
    columns("c").hidden = True
    vla_step = 321 ' vla:1382
    If vlatraceon() Then ' vla:1383
        Call vlatracestep(321, vla_step_text(321)) ' vla:1383
    End If
    columns("c").hidden = False
    vla_step = 322 ' vla:1386
    If vlatraceon() Then ' vla:1387
        Call vlatracestep(322, vla_step_text(322)) ' vla:1387
    End If
    Call columns("b").delete
    vla_step = 323 ' vla:1390
    If vlatraceon() Then ' vla:1391
        Call vlatracestep(323, vla_step_text(323)) ' vla:1391
    End If
    Call rows(5).insert
    vla_step = 324 ' vla:1394
    If vlatraceon() Then ' vla:1395
        Call vlatracestep(324, vla_step_text(324)) ' vla:1395
    End If
    Call rows(1).insert
    vla_step = 325 ' vla:1398
    If vlatraceon() Then ' vla:1399
        Call vlatracestep(325, vla_step_text(325)) ' vla:1399
    End If
    Call rows(2).delete
    vla_step = 326 ' vla:1402
    If vlatraceon() Then ' vla:1403
        Call vlatracestep(326, vla_step_text(326)) ' vla:1403
    End If
    Call rows(3).delete
    vla_step = 327 ' vla:1406
    If vlatraceon() Then ' vla:1407
        Call vlatracestep(327, vla_step_text(327)) ' vla:1407
    End If
    rows(10).hidden = True
    vla_step = 328 ' vla:1410
    If vlatraceon() Then ' vla:1411
        Call vlatracestep(328, vla_step_text(328)) ' vla:1411
    End If
    rows(10).hidden = False
    vla_step = 329 ' vla:1414
    If vlatraceon() Then ' vla:1415
        Call vlatracestep(329, vla_step_text(329)) ' vla:1415
    End If
    Call rows(2).select
    activewindow.freezepanes = True
    vla_step = 330 ' vla:1418
    If vlatraceon() Then ' vla:1419
        Call vlatracestep(330, vla_step_text(330)) ' vla:1419
    End If
    activewindow.freezepanes = False
    vla_step = 331 ' vla:1422
    If vlatraceon() Then ' vla:1423
        Call vlatracestep(331, vla_step_text(331)) ' vla:1423
    End If
    Call columns("a").autofit
    vla_step = 332 ' vla:1426
    If vlatraceon() Then ' vla:1427
        Call vlatracestep(332, vla_step_text(332)) ' vla:1427
    End If
    Call cells.entirecolumn.autofit
    vla_step = 333 ' vla:1430
    If vlatraceon() Then ' vla:1431
        Call vlatracestep(333, vla_step_text(333)) ' vla:1431
    End If
    range("g1") = "Region" ' vla:1433 src:620
    vla_step = 334 ' vla:1434
    If vlatraceon() Then ' vla:1435
        Call vlatracestep(334, vla_step_text(334)) ' vla:1435
    End If
    range("h1") = "Amount" ' vla:1437 src:621
    vla_step = 335 ' vla:1438
    If vlatraceon() Then ' vla:1439
        Call vlatracestep(335, vla_step_text(335)) ' vla:1439
    End If
    range("i1") = "Notes" ' vla:1441 src:622
    vla_step = 336 ' vla:1442
    If vlatraceon() Then ' vla:1443
        Call vlatracestep(336, vla_step_text(336)) ' vla:1443
    End If
    range("g2") = "West" ' vla:1445 src:623
    vla_step = 337 ' vla:1446
    If vlatraceon() Then ' vla:1447
        Call vlatracestep(337, vla_step_text(337)) ' vla:1447
    End If
    range("h2") = 100 ' vla:1449 src:624
    vla_step = 338 ' vla:1450
    If vlatraceon() Then ' vla:1451
        Call vlatracestep(338, vla_step_text(338)) ' vla:1451
    End If
    range("i2") = "ok" ' vla:1453 src:625
    vla_step = 339 ' vla:1454
    If vlatraceon() Then ' vla:1455
        Call vlatracestep(339, vla_step_text(339)) ' vla:1455
    End If
    range("g3") = "East" ' vla:1457 src:626
    vla_step = 340 ' vla:1458
    If vlatraceon() Then ' vla:1459
        Call vlatracestep(340, vla_step_text(340)) ' vla:1459
    End If
    range("h3") = 250 ' vla:1461 src:627
    vla_step = 341 ' vla:1462
    If vlatraceon() Then ' vla:1463
        Call vlatracestep(341, vla_step_text(341)) ' vla:1463
    End If
    range("i3") = "ok" ' vla:1465 src:628
    vla_step = 342 ' vla:1466
    If vlatraceon() Then ' vla:1467
        Call vlatracestep(342, vla_step_text(342)) ' vla:1467
    End If
    range("g4") = "West" ' vla:1469 src:629
    vla_step = 343 ' vla:1470
    If vlatraceon() Then ' vla:1471
        Call vlatracestep(343, vla_step_text(343)) ' vla:1471
    End If
    range("h4") = 100 ' vla:1473 src:630
    vla_step = 344 ' vla:1474
    If vlatraceon() Then ' vla:1475
        Call vlatracestep(344, vla_step_text(344)) ' vla:1475
    End If
    range("i4") = "dup" ' vla:1477 src:631
    vla_step = 345 ' vla:1478
    If vlatraceon() Then ' vla:1479
        Call vlatracestep(345, vla_step_text(345)) ' vla:1479
    End If
    Call range("g1:i4").sort(key1:=range("h1"), order1:=xldescending, header:=xlyes)
    vla_step = 346 ' vla:1482
    If vlatraceon() Then ' vla:1483
        Call vlatracestep(346, vla_step_text(346)) ' vla:1483
    End If
    Call range("g1:i4").autofilter(field:=1, criteria1:="West")
    vla_step = 347 ' vla:1486
    If vlatraceon() Then ' vla:1487
        Call vlatracestep(347, vla_step_text(347)) ' vla:1487
    End If
    On Error Resume Next
    Call activesheet.showalldata
    On Error GoTo 0
    vla_step = 348 ' vla:1490
    If vlatraceon() Then ' vla:1491
        Call vlatracestep(348, vla_step_text(348)) ' vla:1491
    End If
    Call range("g1:i4").replace(what:="dup", replacement:="ok")
    vla_step = 349 ' vla:1494
    If vlatraceon() Then ' vla:1495
        Call vlatracestep(349, vla_step_text(349)) ' vla:1495
    End If
    Call range("g1:i4").removeduplicates(columns:=1, header:=xlyes)
    vla_step = 350 ' vla:1498
    If vlatraceon() Then ' vla:1499
        Call vlatracestep(350, vla_step_text(350)) ' vla:1499
    End If
    Call columns("i").replace(what:="ok", replacement:="fine")
    vla_step = 351 ' vla:1502
    If vlatraceon() Then ' vla:1503
        Call vlatracestep(351, vla_step_text(351)) ' vla:1503
    End If
    Call range("g1:i4").copy
    Call range("g1:i4").pastespecial(paste:=xlpastevalues)
    application.cutcopymode = False
    vla_step = 352 ' vla:1506
    If vlatraceon() Then ' vla:1507
        Call vlatracestep(352, vla_step_text(352)) ' vla:1507
    End If
    Call range("g1:i4").clearformats
    vla_step = 353 ' vla:1510
    If vlatraceon() Then ' vla:1511
        Call vlatracestep(353, vla_step_text(353)) ' vla:1511
    End If
    Call vlacheckrangename("demo_table")
    range("g1:i4").name = "demo_table"
    vla_step = 354 ' vla:1514
    If vlatraceon() Then ' vla:1515
        Call vlatracestep(354, vla_step_text(354)) ' vla:1515
    End If
    Call range("g1:i4").copy(destination:=range("k1:m4"))
    vla_step = 355 ' vla:1518
    If vlatraceon() Then ' vla:1519
        Call vlatracestep(355, vla_step_text(355)) ' vla:1519
    End If
    worksheets("demo").tab.color = vlacolor(hot_pink)
    vla_step = 356 ' vla:1522
    If vlatraceon() Then ' vla:1523
        Call vlatracestep(356, vla_step_text(356)) ' vla:1523
    End If
    Call activesheet.protect(password:="demo123")
    vla_step = 357 ' vla:1526
    If vlatraceon() Then ' vla:1527
        Call vlatracestep(357, vla_step_text(357)) ' vla:1527
    End If
    Call activesheet.unprotect(password:="demo123")
    vla_step = 358 ' vla:1530
    If vlatraceon() Then ' vla:1531
        Call vlatracestep(358, vla_step_text(358)) ' vla:1531
    End If
    On Error GoTo vla_tryf_4 ' vla:1533 src:675
    vla_step = 359 ' vla:1534 src:675
    If vlatraceon() Then ' vla:1535 src:675
        Call vlatracestep(359, vla_step_text(359)) ' vla:1535 src:675
    End If
    application.displayalerts = False
    Call worksheets("gstruct").delete
    application.displayalerts = True
    GoTo vla_tryd_4 ' vla:1538 src:675
vla_tryf_4: ' vla:1539 src:675
    vla_problem = err.description ' vla:1540 src:675
    Resume vla_tryr_4 ' vla:1541 src:675
vla_tryr_4: ' vla:1542 src:675
    On Error GoTo vla_fail ' vla:1543 src:675
vla_tryd_4: ' vla:1544 src:675
    On Error GoTo vla_fail ' vla:1545 src:675
    vla_step = 360 ' vla:1546
    If vlatraceon() Then ' vla:1547
        Call vlatracestep(360, vla_step_text(360)) ' vla:1547
    End If
    Call vlaensuresheet("gstruct") ' vla:1549 src:678
    Call worksheets("gstruct").activate
    vla_step = 361 ' vla:1550
    If vlatraceon() Then ' vla:1551
        Call vlatracestep(361, vla_step_text(361)) ' vla:1551
    End If
    rows(3).hidden = True
    vla_step = 362 ' vla:1554
    If vlatraceon() Then ' vla:1555
        Call vlatracestep(362, vla_step_text(362)) ' vla:1555
    End If
    columns("b").hidden = True
    vla_step = 363 ' vla:1558
    If vlatraceon() Then ' vla:1559
        Call vlatracestep(363, vla_step_text(363)) ' vla:1559
    End If
    cells.entirerow.hidden = False
    cells.entirecolumn.hidden = False
    vla_step = 364 ' vla:1562
    If vlatraceon() Then ' vla:1563
        Call vlatracestep(364, vla_step_text(364)) ' vla:1563
    End If
    range("a6").font.size = 36
    vla_step = 365 ' vla:1566
    If vlatraceon() Then ' vla:1567
        Call vlatracestep(365, vla_step_text(365)) ' vla:1567
    End If
    Call rows(6).autofit
    vla_step = 366 ' vla:1570
    If vlatraceon() Then ' vla:1571
        Call vlatracestep(366, vla_step_text(366)) ' vla:1571
    End If
    Call rows((10 & ":" & 12)).group
    vla_step = 367 ' vla:1574
    If vlatraceon() Then ' vla:1575
        Call vlatracestep(367, vla_step_text(367)) ' vla:1575
    End If
    Call rows((14 & ":" & 16)).group
    vla_step = 368 ' vla:1578
    If vlatraceon() Then ' vla:1579
        Call vlatracestep(368, vla_step_text(368)) ' vla:1579
    End If
    Call rows((14 & ":" & 16)).ungroup
    vla_step = 369 ' vla:1582
    If vlatraceon() Then ' vla:1583
        Call vlatracestep(369, vla_step_text(369)) ' vla:1583
    End If
    range("a20") = "before-insert" ' vla:1585 src:691
    vla_step = 370 ' vla:1586
    If vlatraceon() Then ' vla:1587
        Call vlatracestep(370, vla_step_text(370)) ' vla:1587
    End If
    Call rows(20).resize(rowsize:=3).insert
    vla_step = 371 ' vla:1590
    If vlatraceon() Then ' vla:1591
        Call vlatracestep(371, vla_step_text(371)) ' vla:1591
    End If
    range("a40") = "before-delete" ' vla:1593 src:694
    vla_step = 372 ' vla:1594
    If vlatraceon() Then ' vla:1595
        Call vlatracestep(372, vla_step_text(372)) ' vla:1595
    End If
    Call rows((38 & ":" & 39)).delete
    vla_step = 373 ' vla:1598
    If vlatraceon() Then ' vla:1599
        Call vlatracestep(373, vla_step_text(373)) ' vla:1599
    End If
    range("b50") = "marker-b" ' vla:1601 src:697
    vla_step = 374 ' vla:1602
    If vlatraceon() Then ' vla:1603
        Call vlatracestep(374, vla_step_text(374)) ' vla:1603
    End If
    range("c50") = "marker-c" ' vla:1605 src:698
    vla_step = 375 ' vla:1606
    If vlatraceon() Then ' vla:1607
        Call vlatracestep(375, vla_step_text(375)) ' vla:1607
    End If
    range("d50") = "marker-d" ' vla:1609 src:699
    vla_step = 376 ' vla:1610
    If vlatraceon() Then ' vla:1611
        Call vlatracestep(376, vla_step_text(376)) ' vla:1611
    End If
    Call vlamovecolumn("b", "d")
    vla_step = 377 ' vla:1614
    If vlatraceon() Then ' vla:1615
        Call vlatracestep(377, vla_step_text(377)) ' vla:1615
    End If
    Call vlafreezepanes(2)
    vla_step = 378 ' vla:1618
    If vlatraceon() Then ' vla:1619
        Call vlatracestep(378, vla_step_text(378)) ' vla:1619
    End If
    range("a60").font.bold = True
    vla_step = 379 ' vla:1622
    If vlatraceon() Then ' vla:1623
        Call vlatracestep(379, vla_step_text(379)) ' vla:1623
    End If
    range("a60").interior.color = vbred
    vla_step = 380 ' vla:1626
    If vlatraceon() Then ' vla:1627
        Call vlatracestep(380, vla_step_text(380)) ' vla:1627
    End If
    Call range("a60:a60").copy
    Call range("c60:c60").pastespecial(paste:=xlpasteformats)
    application.cutcopymode = False
    vla_step = 381 ' vla:1630
    If vlatraceon() Then ' vla:1631
        Call vlatracestep(381, vla_step_text(381)) ' vla:1631
    End If
    Call range("a60:a60").copy
    Call range("c62:c62").pastespecial(paste:=xlpasteformats)
    application.cutcopymode = False
    vla_step = 382 ' vla:1634
    If vlatraceon() Then ' vla:1635
        Call vlatracestep(382, vla_step_text(382)) ' vla:1635
    End If
    range("b64").Formula2 = "=5*2"
    vla_step = 383 ' vla:1638
    If vlatraceon() Then ' vla:1639
        Call vlatracestep(383, vla_step_text(383)) ' vla:1639
    End If
    Call range("b64").copy
    Call range("d64").pastespecial(paste:=xlpasteformulas)
    application.cutcopymode = False
    vla_step = 384 ' vla:1642
    If vlatraceon() Then ' vla:1643
        Call vlatracestep(384, vla_step_text(384)) ' vla:1643
    End If
    columns("f").columnwidth = 33
    vla_step = 385 ' vla:1646
    If vlatraceon() Then ' vla:1647
        Call vlatracestep(385, vla_step_text(385)) ' vla:1647
    End If
    Call range("f1:f1").copy
    Call range("h1:h1").pastespecial(paste:=xlpastecolumnwidths)
    application.cutcopymode = False
    vla_step = 386 ' vla:1650
    If vlatraceon() Then ' vla:1651
        Call vlatracestep(386, vla_step_text(386)) ' vla:1651
    End If
    range("a68") = 1 ' vla:1653 src:716
    vla_step = 387 ' vla:1654
    If vlatraceon() Then ' vla:1655
        Call vlatracestep(387, vla_step_text(387)) ' vla:1655
    End If
    range("a69") = 2 ' vla:1657 src:717
    vla_step = 388 ' vla:1658
    If vlatraceon() Then ' vla:1659
        Call vlatracestep(388, vla_step_text(388)) ' vla:1659
    End If
    range("a70") = 3 ' vla:1661 src:718
    vla_step = 389 ' vla:1662
    If vlatraceon() Then ' vla:1663
        Call vlatracestep(389, vla_step_text(389)) ' vla:1663
    End If
    Call range("a68:a70").copy
    Call range("c68").pastespecial(transpose:=True)
    application.cutcopymode = False
    vla_step = 390 ' vla:1666
    If vlatraceon() Then ' vla:1667
        Call vlatracestep(390, vla_step_text(390)) ' vla:1667
    End If
    range("a72") = "cutme" ' vla:1669 src:721
    vla_step = 391 ' vla:1670
    If vlatraceon() Then ' vla:1671
        Call vlatracestep(391, vla_step_text(391)) ' vla:1671
    End If
    Call range("a72:a72").cut(destination:=range("c72"))
    vla_step = 392 ' vla:1674
    If vlatraceon() Then ' vla:1675
        Call vlatracestep(392, vla_step_text(392)) ' vla:1675
    End If
    range("a74") = "rowdata" ' vla:1677 src:724
    vla_step = 393 ' vla:1678
    If vlatraceon() Then ' vla:1679
        Call vlatracestep(393, vla_step_text(393)) ' vla:1679
    End If
    Call rows(74).copy(destination:=rows(76))
    vla_step = 394 ' vla:1682
    If vlatraceon() Then ' vla:1683
        Call vlatracestep(394, vla_step_text(394)) ' vla:1683
    End If
    range("a80") = "clearme" ' vla:1685 src:728
    vla_step = 395 ' vla:1686
    If vlatraceon() Then ' vla:1687
        Call vlatracestep(395, vla_step_text(395)) ' vla:1687
    End If
    range("a80").font.bold = True
    vla_step = 396 ' vla:1690
    If vlatraceon() Then ' vla:1691
        Call vlatracestep(396, vla_step_text(396)) ' vla:1691
    End If
    range("a80").interior.color = vbred
    vla_step = 397 ' vla:1694
    If vlatraceon() Then ' vla:1695
        Call vlatracestep(397, vla_step_text(397)) ' vla:1695
    End If
    Call range("a80:b80").clear
    vla_step = 398 ' vla:1698
    If vlatraceon() Then ' vla:1699
        Call vlatracestep(398, vla_step_text(398)) ' vla:1699
    End If
    range("a84") = "m1" ' vla:1701 src:733
    vla_step = 399 ' vla:1702
    If vlatraceon() Then ' vla:1703
        Call vlatracestep(399, vla_step_text(399)) ' vla:1703
    End If
    range("a85") = "m2" ' vla:1705 src:734
    vla_step = 400 ' vla:1706
    If vlatraceon() Then ' vla:1707
        Call vlatracestep(400, vla_step_text(400)) ' vla:1707
    End If
    range("a86") = "m3" ' vla:1709 src:735
    vla_step = 401 ' vla:1710
    If vlatraceon() Then ' vla:1711
        Call vlatracestep(401, vla_step_text(401)) ' vla:1711
    End If
    range("a87") = "m4" ' vla:1713 src:736
    vla_step = 402 ' vla:1714
    If vlatraceon() Then ' vla:1715
        Call vlatracestep(402, vla_step_text(402)) ' vla:1715
    End If
    range("a88") = "m5" ' vla:1717 src:737
    vla_step = 403 ' vla:1718
    If vlatraceon() Then ' vla:1719
        Call vlatracestep(403, vla_step_text(403)) ' vla:1719
    End If
    Call range("a84:a86").delete(shift:=xlshiftup)
    vla_step = 404 ' vla:1722
    If vlatraceon() Then ' vla:1723
        Call vlatracestep(404, vla_step_text(404)) ' vla:1723
    End If
    range("b90") = "x1" ' vla:1725 src:740
    vla_step = 405 ' vla:1726
    If vlatraceon() Then ' vla:1727
        Call vlatracestep(405, vla_step_text(405)) ' vla:1727
    End If
    range("c90") = "x2" ' vla:1729 src:741
    vla_step = 406 ' vla:1730
    If vlatraceon() Then ' vla:1731
        Call vlatracestep(406, vla_step_text(406)) ' vla:1731
    End If
    range("d90") = "x3" ' vla:1733 src:742
    vla_step = 407 ' vla:1734
    If vlatraceon() Then ' vla:1735
        Call vlatracestep(407, vla_step_text(407)) ' vla:1735
    End If
    range("e90") = "rightdata" ' vla:1737 src:743
    vla_step = 408 ' vla:1738
    If vlatraceon() Then ' vla:1739
        Call vlatracestep(408, vla_step_text(408)) ' vla:1739
    End If
    Call range("b90:d90").delete(shift:=xlshifttoleft)
    vla_step = 409 ' vla:1742
    If vlatraceon() Then ' vla:1743
        Call vlatracestep(409, vla_step_text(409)) ' vla:1743
    End If
    range("a94") = 1 ' vla:1745 src:746
    vla_step = 410 ' vla:1746
    If vlatraceon() Then ' vla:1747
        Call vlatracestep(410, vla_step_text(410)) ' vla:1747
    End If
    range("b94") = 2 ' vla:1749 src:747
    vla_step = 411 ' vla:1750
    If vlatraceon() Then ' vla:1751
        Call vlatracestep(411, vla_step_text(411)) ' vla:1751
    End If
    range("a95") = 3 ' vla:1753 src:748
    vla_step = 412 ' vla:1754
    If vlatraceon() Then ' vla:1755
        Call vlatracestep(412, vla_step_text(412)) ' vla:1755
    End If
    range("a96") = 5 ' vla:1757 src:749
    vla_step = 413 ' vla:1758
    If vlatraceon() Then ' vla:1759
        Call vlatracestep(413, vla_step_text(413)) ' vla:1759
    End If
    range("b96") = 6 ' vla:1761 src:750
    vla_step = 414 ' vla:1762
    If vlatraceon() Then ' vla:1763
        Call vlatracestep(414, vla_step_text(414)) ' vla:1763
    End If
    Call vladeleteblankrows(range("a94:b96"))
    vla_step = 415 ' vla:1766
    If vlatraceon() Then ' vla:1767
        Call vlatracestep(415, vla_step_text(415)) ' vla:1767
    End If
    range("a150") = 1 ' vla:1769 src:754
    vla_step = 416 ' vla:1770
    If vlatraceon() Then ' vla:1771
        Call vlatracestep(416, vla_step_text(416)) ' vla:1771
    End If
    range("a151") = 2 ' vla:1773 src:755
    vla_step = 417 ' vla:1774
    If vlatraceon() Then ' vla:1775
        Call vlatracestep(417, vla_step_text(417)) ' vla:1775
    End If
    range("a152") = 3 ' vla:1777 src:756
    vla_step = 418 ' vla:1778
    If vlatraceon() Then ' vla:1779
        Call vlatracestep(418, vla_step_text(418)) ' vla:1779
    End If
    range("a153") = 4 ' vla:1781 src:757
    vla_step = 419 ' vla:1782
    If vlatraceon() Then ' vla:1783
        Call vlatracestep(419, vla_step_text(419)) ' vla:1783
    End If
    Call vlabandrows(range("a150:d153"), vlacolor("#D9D9D9"))
    vla_step = 420 ' vla:1786
    If vlatraceon() Then ' vla:1787
        Call vlatracestep(420, vla_step_text(420)) ' vla:1787
    End If
    rows(165).font.bold = True
    rows(165).interior.color = vlacolor("#D9D9D9")
    vla_step = 421 ' vla:1790
    If vlatraceon() Then ' vla:1791
        Call vlatracestep(421, vla_step_text(421)) ' vla:1791
    End If
    Call vlafillseries(range("a170:a179"), 1, 1, "linear")
    vla_step = 422 ' vla:1794
    If vlatraceon() Then ' vla:1795
        Call vlatracestep(422, vla_step_text(422)) ' vla:1795
    End If
    Call vlafillseries(range("a180:a184"), 1, 2, "linear")
    vla_step = 423 ' vla:1798
    If vlatraceon() Then ' vla:1799
        Call vlatracestep(423, vla_step_text(423)) ' vla:1799
    End If
    Call vlafillseries(range("a190:a194"), 2, 2, "growth")
    vla_step = 424 ' vla:1802
    If vlatraceon() Then ' vla:1803
        Call vlatracestep(424, vla_step_text(424)) ' vla:1803
    End If
    Call vlafillseries(range("a200:a203"), 2, 3, "growth")
    vla_step = 425 ' vla:1806
    If vlatraceon() Then ' vla:1807
        Call vlatracestep(425, vla_step_text(425)) ' vla:1807
    End If
    range("z500").font.bold = True
    vla_step = 426 ' vla:1810
    If vlatraceon() Then ' vla:1811
        Call vlatracestep(426, vla_step_text(426)) ' vla:1811
    End If
    Call range("z500").clear
    vla_step = 427 ' vla:1814
    If vlatraceon() Then ' vla:1815
        Call vlatracestep(427, vla_step_text(427)) ' vla:1815
    End If
    Call vlatrimsheet
    vla_step = 428 ' vla:1818
    If vlatraceon() Then ' vla:1819
        Call vlatracestep(428, vla_step_text(428)) ' vla:1819
    End If
    On Error GoTo vla_tryf_5 ' vla:1821 src:792
    vla_step = 429 ' vla:1822 src:792
    If vlatraceon() Then ' vla:1823 src:792
        Call vlatracestep(429, vla_step_text(429)) ' vla:1823 src:792
    End If
    application.displayalerts = False
    Call worksheets("gformat").delete
    application.displayalerts = True
    GoTo vla_tryd_5 ' vla:1826 src:792
vla_tryf_5: ' vla:1827 src:792
    vla_problem = err.description ' vla:1828 src:792
    Resume vla_tryr_5 ' vla:1829 src:792
vla_tryr_5: ' vla:1830 src:792
    On Error GoTo vla_fail ' vla:1831 src:792
vla_tryd_5: ' vla:1832 src:792
    On Error GoTo vla_fail ' vla:1833 src:792
    vla_step = 430 ' vla:1834
    If vlatraceon() Then ' vla:1835
        Call vlatracestep(430, vla_step_text(430)) ' vla:1835
    End If
    Call vlaensuresheet("gformat") ' vla:1837 src:795
    Call worksheets("gformat").activate
    vla_step = 431 ' vla:1838
    If vlatraceon() Then ' vla:1839
        Call vlatracestep(431, vla_step_text(431)) ' vla:1839
    End If
    range("a1:c1").font.bold = True
    vla_step = 432 ' vla:1842
    If vlatraceon() Then ' vla:1843
        Call vlatracestep(432, vla_step_text(432)) ' vla:1843
    End If
    range("a2:c2").font.italic = True
    vla_step = 433 ' vla:1846
    If vlatraceon() Then ' vla:1847
        Call vlatracestep(433, vla_step_text(433)) ' vla:1847
    End If
    range("a3:c3").interior.color = vbblue
    vla_step = 434 ' vla:1850
    If vlatraceon() Then ' vla:1851
        Call vlatracestep(434, vla_step_text(434)) ' vla:1851
    End If
    range("a4:c4").font.color = vlacolor("#FF0000")
    vla_step = 435 ' vla:1854
    If vlatraceon() Then ' vla:1855
        Call vlatracestep(435, vla_step_text(435)) ' vla:1855
    End If
    range("a5:c5").interior.color = vlacolor("#00FF00")
    vla_step = 436 ' vla:1858
    If vlatraceon() Then ' vla:1859
        Call vlatracestep(436, vla_step_text(436)) ' vla:1859
    End If
    range("a6:c6").interior.color = vbyellow
    vla_step = 437 ' vla:1862
    If vlatraceon() Then ' vla:1863
        Call vlatracestep(437, vla_step_text(437)) ' vla:1863
    End If
    range("b6:c6").interior.colorindex = xlnone
    vla_step = 438 ' vla:1866
    If vlatraceon() Then ' vla:1867
        Call vlatracestep(438, vla_step_text(438)) ' vla:1867
    End If
    range("a7:c7").font.size = 16
    vla_step = 439 ' vla:1870
    If vlatraceon() Then ' vla:1871
        Call vlatracestep(439, vla_step_text(439)) ' vla:1871
    End If
    range("a8:c8").font.bold = True
    vla_step = 440 ' vla:1874
    If vlatraceon() Then ' vla:1875
        Call vlatracestep(440, vla_step_text(440)) ' vla:1875
    End If
    range("b8:c8").font.bold = False
    vla_step = 441 ' vla:1878
    If vlatraceon() Then ' vla:1879
        Call vlatracestep(441, vla_step_text(441)) ' vla:1879
    End If
    range("a9:c9").font.italic = True
    vla_step = 442 ' vla:1882
    If vlatraceon() Then ' vla:1883
        Call vlatracestep(442, vla_step_text(442)) ' vla:1883
    End If
    range("c9").font.italic = False
    vla_step = 443 ' vla:1886
    If vlatraceon() Then ' vla:1887
        Call vlatracestep(443, vla_step_text(443)) ' vla:1887
    End If
    range("a10:c10").font.underline = xlunderlinestylesingle
    vla_step = 444 ' vla:1890
    If vlatraceon() Then ' vla:1891
        Call vlatracestep(444, vla_step_text(444)) ' vla:1891
    End If
    range("a11:c11").font.strikethrough = True
    vla_step = 445 ' vla:1894
    If vlatraceon() Then ' vla:1895
        Call vlatracestep(445, vla_step_text(445)) ' vla:1895
    End If
    range("a12:c12").font.underline = xlunderlinestylesingle
    vla_step = 446 ' vla:1898
    If vlatraceon() Then ' vla:1899
        Call vlatracestep(446, vla_step_text(446)) ' vla:1899
    End If
    range("b12:c12").font.underline = xlunderlinestylenone
    vla_step = 447 ' vla:1902
    If vlatraceon() Then ' vla:1903
        Call vlatracestep(447, vla_step_text(447)) ' vla:1903
    End If
    range("a13:c13").font.strikethrough = True
    vla_step = 448 ' vla:1906
    If vlatraceon() Then ' vla:1907
        Call vlatracestep(448, vla_step_text(448)) ' vla:1907
    End If
    range("c13").font.strikethrough = False
    vla_step = 449 ' vla:1910
    If vlatraceon() Then ' vla:1911
        Call vlatracestep(449, vla_step_text(449)) ' vla:1911
    End If
    range("a14:c14").font.name = "Courier New"
    vla_step = 450 ' vla:1914
    If vlatraceon() Then ' vla:1915
        Call vlatracestep(450, vla_step_text(450)) ' vla:1915
    End If
    range("a15:c15").verticalalignment = xltop
    vla_step = 451 ' vla:1918
    If vlatraceon() Then ' vla:1919
        Call vlatracestep(451, vla_step_text(451)) ' vla:1919
    End If
    range("a16:c16").verticalalignment = xlcenter
    vla_step = 452 ' vla:1922
    If vlatraceon() Then ' vla:1923
        Call vlatracestep(452, vla_step_text(452)) ' vla:1923
    End If
    range("a17:c17").verticalalignment = xltop
    vla_step = 453 ' vla:1926
    If vlatraceon() Then ' vla:1927
        Call vlatracestep(453, vla_step_text(453)) ' vla:1927
    End If
    range("b17:c17").verticalalignment = xlbottom
    vla_step = 454 ' vla:1930
    If vlatraceon() Then ' vla:1931
        Call vlatracestep(454, vla_step_text(454)) ' vla:1931
    End If
    range("a18:c18").indentlevel = 2
    vla_step = 455 ' vla:1934
    If vlatraceon() Then ' vla:1935
        Call vlatracestep(455, vla_step_text(455)) ' vla:1935
    End If
    range("a19:c19").orientation = 45
    vla_step = 456 ' vla:1938
    If vlatraceon() Then ' vla:1939
        Call vlatracestep(456, vla_step_text(456)) ' vla:1939
    End If
    range("b21:d23").borders(xledgetop).linestyle = xlcontinuous
    range("b21:d23").borders(xledgebottom).linestyle = xlcontinuous
    range("b21:d23").borders(xledgeleft).linestyle = xlcontinuous
    range("b21:d23").borders(xledgeright).linestyle = xlcontinuous
    vla_step = 457 ' vla:1942
    If vlatraceon() Then ' vla:1943
        Call vlatracestep(457, vla_step_text(457)) ' vla:1943
    End If
    range("b25:d25").borders(xledgebottom).linestyle = xlcontinuous
    vla_step = 458 ' vla:1946
    If vlatraceon() Then ' vla:1947
        Call vlatracestep(458, vla_step_text(458)) ' vla:1947
    End If
    range("b27:d29").borders.linestyle = xlcontinuous
    vla_step = 459 ' vla:1950
    If vlatraceon() Then ' vla:1951
        Call vlatracestep(459, vla_step_text(459)) ' vla:1951
    End If
    range("b31:d33").borders.linestyle = xlcontinuous
    vla_step = 460 ' vla:1954
    If vlatraceon() Then ' vla:1955
        Call vlatracestep(460, vla_step_text(460)) ' vla:1955
    End If
    range("c31:d33").borders.linestyle = xlnone
    vla_step = 461 ' vla:1958
    If vlatraceon() Then ' vla:1959
        Call vlatracestep(461, vla_step_text(461)) ' vla:1959
    End If
    columns("f").columnwidth = 40
    vla_step = 462 ' vla:1962
    If vlatraceon() Then ' vla:1963
        Call vlatracestep(462, vla_step_text(462)) ' vla:1963
    End If
    range("f40") = 1234.56 ' vla:1965 src:847
    vla_step = 463 ' vla:1966
    If vlatraceon() Then ' vla:1967
        Call vlatracestep(463, vla_step_text(463)) ' vla:1967
    End If
    range("f40").numberformat = vlanumberformatcode("number", 2)
    vla_step = 464 ' vla:1970
    If vlatraceon() Then ' vla:1971
        Call vlatracestep(464, vla_step_text(464)) ' vla:1971
    End If
    range("f41") = 1234.56 ' vla:1973 src:849
    vla_step = 465 ' vla:1974
    If vlatraceon() Then ' vla:1975
        Call vlatracestep(465, vla_step_text(465)) ' vla:1975
    End If
    range("f41").numberformat = vlanumberformatcode("number", 3)
    vla_step = 466 ' vla:1978
    If vlatraceon() Then ' vla:1979
        Call vlatracestep(466, vla_step_text(466)) ' vla:1979
    End If
    range("f42") = 1234.56 ' vla:1981 src:851
    vla_step = 467 ' vla:1982
    If vlatraceon() Then ' vla:1983
        Call vlatracestep(467, vla_step_text(467)) ' vla:1983
    End If
    range("f42").numberformat = vlanumberformatcode("number-separated", 2)
    vla_step = 468 ' vla:1986
    If vlatraceon() Then ' vla:1987
        Call vlatracestep(468, vla_step_text(468)) ' vla:1987
    End If
    range("f43") = 1234.56 ' vla:1989 src:853
    vla_step = 469 ' vla:1990
    If vlatraceon() Then ' vla:1991
        Call vlatracestep(469, vla_step_text(469)) ' vla:1991
    End If
    range("f43").numberformat = vlanumberformatcode("number-separated", 0)
    vla_step = 470 ' vla:1994
    If vlatraceon() Then ' vla:1995
        Call vlatracestep(470, vla_step_text(470)) ' vla:1995
    End If
    range("f44") = 1234.56 ' vla:1997 src:855
    vla_step = 471 ' vla:1998
    If vlatraceon() Then ' vla:1999
        Call vlatracestep(471, vla_step_text(471)) ' vla:1999
    End If
    range("f44").numberformat = vlanumberformatcode("dollars", 2)
    vla_step = 472 ' vla:2002
    If vlatraceon() Then ' vla:2003
        Call vlatracestep(472, vla_step_text(472)) ' vla:2003
    End If
    range("f45") = 1234.56 ' vla:2005 src:857
    vla_step = 473 ' vla:2006
    If vlatraceon() Then ' vla:2007
        Call vlatracestep(473, vla_step_text(473)) ' vla:2007
    End If
    range("f45").numberformat = vlanumberformatcode("euros", 0)
    vla_step = 474 ' vla:2010
    If vlatraceon() Then ' vla:2011
        Call vlatracestep(474, vla_step_text(474)) ' vla:2011
    End If
    range("f46") = 1234.56 ' vla:2013 src:859
    vla_step = 475 ' vla:2014
    If vlatraceon() Then ' vla:2015
        Call vlatracestep(475, vla_step_text(475)) ' vla:2015
    End If
    range("f46").numberformat = vlanumberformatcode("pounds", 2)
    vla_step = 476 ' vla:2018
    If vlatraceon() Then ' vla:2019
        Call vlatracestep(476, vla_step_text(476)) ' vla:2019
    End If
    range("f47") = 1234.56 ' vla:2021 src:861
    vla_step = 477 ' vla:2022
    If vlatraceon() Then ' vla:2023
        Call vlatracestep(477, vla_step_text(477)) ' vla:2023
    End If
    range("f47").numberformat = vlanumberformatcode("accounting-dollars", 2)
    vla_step = 478 ' vla:2026
    If vlatraceon() Then ' vla:2027
        Call vlatracestep(478, vla_step_text(478)) ' vla:2027
    End If
    range("f48") = -1234.56 ' vla:2029 src:863
    vla_step = 479 ' vla:2030
    If vlatraceon() Then ' vla:2031
        Call vlatracestep(479, vla_step_text(479)) ' vla:2031
    End If
    range("f48").numberformat = vlanumberformatcode("accounting-euros", 0)
    vla_step = 480 ' vla:2034
    If vlatraceon() Then ' vla:2035
        Call vlatracestep(480, vla_step_text(480)) ' vla:2035
    End If
    range("f49") = 0.125 ' vla:2037 src:865
    vla_step = 481 ' vla:2038
    If vlatraceon() Then ' vla:2039
        Call vlatracestep(481, vla_step_text(481)) ' vla:2039
    End If
    range("f49").numberformat = "0.0%"
    vla_step = 482 ' vla:2042
    If vlatraceon() Then ' vla:2043
        Call vlatracestep(482, vla_step_text(482)) ' vla:2043
    End If
    range("f50") = 0.125 ' vla:2045 src:867
    vla_step = 483 ' vla:2046
    If vlatraceon() Then ' vla:2047
        Call vlatracestep(483, vla_step_text(483)) ' vla:2047
    End If
    range("f50").numberformat = vlanumberformatcode("percent", 2)
    vla_step = 484 ' vla:2050
    If vlatraceon() Then ' vla:2051
        Call vlatracestep(484, vla_step_text(484)) ' vla:2051
    End If
    range("f51") = 46000 ' vla:2053 src:869
    vla_step = 485 ' vla:2054
    If vlatraceon() Then ' vla:2055
        Call vlatracestep(485, vla_step_text(485)) ' vla:2055
    End If
    range("f51").numberformat = "m/d/yyyy"
    vla_step = 486 ' vla:2058
    If vlatraceon() Then ' vla:2059
        Call vlatracestep(486, vla_step_text(486)) ' vla:2059
    End If
    range("f52") = 46000 ' vla:2061 src:871
    vla_step = 487 ' vla:2062
    If vlatraceon() Then ' vla:2063
        Call vlatracestep(487, vla_step_text(487)) ' vla:2063
    End If
    range("f52").numberformat = "[$-F800]dddd, mmmm dd, yyyy"
    vla_step = 488 ' vla:2066
    If vlatraceon() Then ' vla:2067
        Call vlatracestep(488, vla_step_text(488)) ' vla:2067
    End If
    range("f53") = 46000 ' vla:2069 src:873
    vla_step = 489 ' vla:2070
    If vlatraceon() Then ' vla:2071
        Call vlatracestep(489, vla_step_text(489)) ' vla:2071
    End If
    range("f53").numberformat = "yyyy-mm-dd"
    vla_step = 490 ' vla:2074
    If vlatraceon() Then ' vla:2075
        Call vlatracestep(490, vla_step_text(490)) ' vla:2075
    End If
    range("f54") = 0.5625 ' vla:2077 src:875
    vla_step = 491 ' vla:2078
    If vlatraceon() Then ' vla:2079
        Call vlatracestep(491, vla_step_text(491)) ' vla:2079
    End If
    range("f54").numberformat = "[$-F400]h:mm:ss AM/PM"
    vla_step = 492 ' vla:2082
    If vlatraceon() Then ' vla:2083
        Call vlatracestep(492, vla_step_text(492)) ' vla:2083
    End If
    range("f55") = 42 ' vla:2085 src:877
    vla_step = 493 ' vla:2086
    If vlatraceon() Then ' vla:2087
        Call vlatracestep(493, vla_step_text(493)) ' vla:2087
    End If
    range("f55").numberformat = "@"
    vla_step = 494 ' vla:2090
    If vlatraceon() Then ' vla:2091
        Call vlatracestep(494, vla_step_text(494)) ' vla:2091
    End If
    range("f56") = 1234.56 ' vla:2093 src:879
    vla_step = 495 ' vla:2094
    If vlatraceon() Then ' vla:2095
        Call vlatracestep(495, vla_step_text(495)) ' vla:2095
    End If
    range("f56").numberformat = vlanumberformatcode("dollars", 2)
    vla_step = 496 ' vla:2098
    If vlatraceon() Then ' vla:2099
        Call vlatracestep(496, vla_step_text(496)) ' vla:2099
    End If
    range("f56").numberformat = "General"
    vla_step = 497 ' vla:2102
    If vlatraceon() Then ' vla:2103
        Call vlatracestep(497, vla_step_text(497)) ' vla:2103
    End If
    range("f57") = 42 ' vla:2105 src:882
    vla_step = 498 ' vla:2106
    If vlatraceon() Then ' vla:2107
        Call vlatracestep(498, vla_step_text(498)) ' vla:2107
    End If
    range("f57").numberformat = "00000"
    vla_step = 499 ' vla:2110
    If vlatraceon() Then ' vla:2111
        Call vlatracestep(499, vla_step_text(499)) ' vla:2111
    End If
    range("h40:j42").borders(xledgetop).linestyle = xlcontinuous
    range("h40:j42").borders(xledgetop).color = vlacolor("#FF0000")
    range("h40:j42").borders(xledgebottom).linestyle = xlcontinuous
    range("h40:j42").borders(xledgebottom).color = vlacolor("#FF0000")
    range("h40:j42").borders(xledgeleft).linestyle = xlcontinuous
    range("h40:j42").borders(xledgeleft).color = vlacolor("#FF0000")
    range("h40:j42").borders(xledgeright).linestyle = xlcontinuous
    range("h40:j42").borders(xledgeright).color = vlacolor("#FF0000")
    vla_step = 500 ' vla:2114
    If vlatraceon() Then ' vla:2115
        Call vlatracestep(500, vla_step_text(500)) ' vla:2115
    End If
    range("h44:j44").borders(xledgebottom).linestyle = xlcontinuous
    range("h44:j44").borders(xledgebottom).color = vlacolor("green")
    vla_step = 501 ' vla:2118
    If vlatraceon() Then ' vla:2119
        Call vlatracestep(501, vla_step_text(501)) ' vla:2119
    End If
    range("h46:j48").borders.linestyle = xlcontinuous
    range("h46:j48").borders.color = vlacolor("blue")
    vla_step = 502 ' vla:2122
    If vlatraceon() Then ' vla:2123
        Call vlatracestep(502, vla_step_text(502)) ' vla:2123
    End If
    Call worksheets("output").activate
    vla_step = 503 ' vla:2126
    If vlatraceon() Then ' vla:2127
        Call vlatracestep(503, vla_step_text(503)) ' vla:2127
    End If
    Call tidy_up ' vla:2129 src:895
    vla_step = 504 ' vla:2130
    If vlatraceon() Then ' vla:2131
        Call vlatracestep(504, vla_step_text(504)) ' vla:2131
    End If
    application.screenupdating = True
    vla_step = 505 ' vla:2134
    If vlatraceon() Then ' vla:2135
        Call vlatracestep(505, vla_step_text(505)) ' vla:2135
    End If
    Debug.Print "report finished" ' vla:2137 src:897
    Exit Sub ' vla:2138
vla_fail: ' vla:2139
    Call vla_report_error ' vla:2140
End Sub

Public Sub vla_report_error()
    Call vlashowerror(("Something went wrong at step " & vla_step & ":" & vbcrlf & vbcrlf & vla_step_text(vla_step) & vbcrlf & vbcrlf & "Excel says: " & err.description)) ' vla:2143
End Sub

Public Function vla_step_text(ByVal n As Long) As String
    Select Case n ' vla:2148
        Case 1
            vla_step_text = "Work on sheet Output. [line 18]" ' vla:2149
            Exit Function
        Case 2
            vla_step_text = "Turn off screen updating. [line 20]" ' vla:2150
            Exit Function
        Case 3
            vla_step_text = "Create a number called total. [line 21]" ' vla:2151
            Exit Function
        Case 4
            vla_step_text = "Set total to 0. [line 22]" ' vla:2152
            Exit Function
        Case 5
            vla_step_text = "Put ""Test Report"" into cell A1. [line 23]" ' vla:2153
            Exit Function
        Case 6
            vla_step_text = "Make cell A1 bold. [line 24]" ' vla:2154
            Exit Function
        Case 7
            vla_step_text = "Set font size of cell A1 to 14. [line 25]" ' vla:2155
            Exit Function
        Case 8
            vla_step_text = "Put today into cell D1. [line 26]" ' vla:2156
            Exit Function
        Case 9
            vla_step_text = "Repeat 5 times: [line 28]" ' vla:2157
            Exit Function
        Case 10
            vla_step_text = "Increase total by counter. [line 29]" ' vla:2158
            Exit Function
        Case 11
            vla_step_text = "If counter is divisible by 2, log ""even step "" joined with counter. [line 30]" ' vla:2159
            Exit Function
        Case 12
            vla_step_text = "Put total into cell B2. [line 32]" ' vla:2160
            Exit Function
        Case 13
            vla_step_text = "Put formula ""=B2*2"" into cell B3. [line 33]" ' vla:2161
            Exit Function
        Case 14
            vla_step_text = "Put sum of range ""B2:B3"" into cell B4. [line 37]" ' vla:2162
            Exit Function
        Case 15
            vla_step_text = "Set grand to sum of range B2:B3. [line 38]" ' vla:2163
            Exit Function
        Case 16
            vla_step_text = "Log ""grand is "" joined with grand. [line 39]" ' vla:2164
            Exit Function
        Case 17
            vla_step_text = "If grand is greater than 40: [line 43]" ' vla:2165
            Exit Function
        Case 18
            vla_step_text = "Put ""PASS"" into cell C4. [line 44]" ' vla:2166
            Exit Function
        Case 19
            vla_step_text = "Repeat 2 times: [line 45]" ' vla:2167
            Exit Function
        Case 20
            vla_step_text = "Log ""pass check "" joined with counter. [line 46]" ' vla:2168
            Exit Function
        Case 21
            vla_step_text = "Make cell C4 bold. [line 48]" ' vla:2169
            Exit Function
        Case 22
            vla_step_text = "Put ""CHECK"" into cell C4. [line 51]" ' vla:2170
            Exit Function
        Case 23
            vla_step_text = "If grand is at least 45, make cell C4 yellow. [line 53]" ' vla:2171
            Exit Function
        Case 24
            vla_step_text = "Create a text called label. [line 54]" ' vla:2172
            Exit Function
        Case 25
            vla_step_text = "Set label to ""Total: "" joined with total. [line 55]" ' vla:2173
            Exit Function
        Case 26
            vla_step_text = "Put label into cell A6. [line 56]" ' vla:2174
            Exit Function
        Case 27
            vla_step_text = "Set font-color of cell A6 to hot-pink. [line 57]" ' vla:2175
            Exit Function
        Case 28
            vla_step_text = "Remember range B2:B4 as results. [line 58]" ' vla:2176
            Exit Function
        Case 29
            vla_step_text = "For each r in results, log r. [line 59]" ' vla:2177
            Exit Function
        Case 30
            vla_step_text = "Set biggest to largest of results. [line 60]" ' vla:2178
            Exit Function
        Case 31
            vla_step_text = "Log ""largest result is "" joined with biggest joined with "", label length "" joined with length of label. [line 61]" ' vla:2179
            Exit Function
        Case 32
            vla_step_text = "Repeat 3 times: [line 64]" ' vla:2180
            Exit Function
        Case 33
            vla_step_text = "Put counter times 10 into column E row counter. [line 65]" ' vla:2181
            Exit Function
        Case 34
            vla_step_text = "Set probe to cell in column E row 2. [line 67]" ' vla:2182
            Exit Function
        Case 35
            vla_step_text = "Create a text called num-col-check. [line 68]" ' vla:2183
            Exit Function
        Case 36
            vla_step_text = "If cell in column number 5 row 2 is 20, set num-col-check to ""yes"". [line 69]" ' vla:2184
            Exit Function
        Case 37
            vla_step_text = "Create a text called value-word-check. [line 70]" ' vla:2185
            Exit Function
        Case 38
            vla_step_text = "If value in column number 5 row 3 is 30, set value-word-check to ""ok"". [line 71]" ' vla:2186
            Exit Function
        Case 39
            vla_step_text = "Put ""bang"" into cell Output!H16. [line 72]" ' vla:2187
            Exit Function
        Case 40
            vla_step_text = "Center cell A1. [line 73]" ' vla:2188
            Exit Function
        Case 41
            vla_step_text = "Set width of column D to 24. [line 74]" ' vla:2189
            Exit Function
        Case 42
            vla_step_text = "Set round-check to 3.14159 rounded to 2 decimals. [line 75]" ' vla:2190
            Exit Function
        Case 43
            vla_step_text = "Log ""rounded is "" joined with round-check. [line 76]" ' vla:2191
            Exit Function
        Case 44
            vla_step_text = "Set thousand-check to 1,000 plus 500. [line 77]" ' vla:2192
            Exit Function
        Case 45
            vla_step_text = "Log ""thousands read as "" joined with thousand-check. [line 78]" ' vla:2193
            Exit Function
        Case 46
            vla_step_text = "Log ""probe is "" joined with probe. [line 79]" ' vla:2194
            Exit Function
        Case 47
            vla_step_text = "Create a number called countdown. [line 81]" ' vla:2195
            Exit Function
        Case 48
            vla_step_text = "Set countdown to 3. [line 82]" ' vla:2196
            Exit Function
        Case 49
            vla_step_text = "While countdown is greater than 0: [line 83]" ' vla:2197
            Exit Function
        Case 50
            vla_step_text = "Log ""countdown "" joined with countdown. [line 84]" ' vla:2198
            Exit Function
        Case 51
            vla_step_text = "Decrease countdown by 1. [line 85]" ' vla:2199
            Exit Function
        Case 52
            vla_step_text = "Put value into column F row row-number. [line 90]" ' vla:2200
            Exit Function
        Case 53
            vla_step_text = "Stamp. [line 92]" ' vla:2201
            Exit Function
        Case 54
            vla_step_text = "Stamp with row-number of 2 and value of ""beta"". [line 93]" ' vla:2202
            Exit Function
        Case 55
            vla_step_text = "Set f-last to last filled row of column F. [line 94]" ' vla:2203
            Exit Function
        Case 56
            vla_step_text = "Log ""column F filled to row "" joined with f-last. [line 95]" ' vla:2204
            Exit Function
        Case 57
            vla_step_text = "Count echo-row from 1 to f-last, log ""echo "" joined with echo-row. [line 96]" ' vla:2205
            Exit Function
        Case 58
            vla_step_text = "Count check-row from 1 to f-last: [line 97]" ' vla:2206
            Exit Function
        Case 59
            vla_step_text = "If cell in column F row check-row contains ""ok"", make cell in column F row check-row bold. [line 98]" ' vla:2207
            Exit Function
        Case 60
            vla_step_text = "Count stripe-row from 1 to f-last step 2: [line 103]" ' vla:2208
            Exit Function
        Case 61
            vla_step_text = "Make cell in column F row stripe-row bold. [line 104]" ' vla:2209
            Exit Function
        Case 62
            vla_step_text = "Count back-row down from f-last to 1 step 2: [line 106]" ' vla:2210
            Exit Function
        Case 63
            vla_step_text = "Log ""back-row "" joined with back-row. [line 107]" ' vla:2211
            Exit Function
        Case 64
            vla_step_text = "Count search-row from 1 to 10: [line 111]" ' vla:2212
            Exit Function
        Case 65
            vla_step_text = "If search-row is 3, stop the loop. [line 112]" ' vla:2213
            Exit Function
        Case 66
            vla_step_text = "Log ""stopped at "" joined with search-row. [line 114]" ' vla:2214
            Exit Function
        Case 67
            vla_step_text = "Create a list called found-items. [line 119]" ' vla:2215
            Exit Function
        Case 68
            vla_step_text = "Repeat 4 times: [line 120]" ' vla:2216
            Exit Function
        Case 69
            vla_step_text = "If counter is greater than 2, append counter times 100 to found-items. [line 121]" ' vla:2217
            Exit Function
        Case 70
            vla_step_text = "For each f in found-items, log ""found "" joined with f. [line 123]" ' vla:2218
            Exit Function
        Case 71
            vla_step_text = "Set list-count to count of found-items. [line 124]" ' vla:2219
            Exit Function
        Case 72
            vla_step_text = "Log ""list holds "" joined with list-count. [line 125]" ' vla:2220
            Exit Function
        Case 73
            vla_step_text = "Create a text called verdict. [line 131]" ' vla:2221
            Exit Function
        Case 74
            vla_step_text = "If grand is greater than 100: [line 132]" ' vla:2222
            Exit Function
        Case 75
            vla_step_text = "Set verdict to ""huge"". [line 133]" ' vla:2223
            Exit Function
        Case 76
            vla_step_text = "Set verdict to ""solid"". [line 136]" ' vla:2224
            Exit Function
        Case 77
            vla_step_text = "Set verdict to ""small"". [line 139]" ' vla:2225
            Exit Function
        Case 78
            vla_step_text = "Create a text called region-label. [line 141]" ' vla:2226
            Exit Function
        Case 79
            vla_step_text = "Set region to ""South"". [line 142]" ' vla:2227
            Exit Function
        Case 80
            vla_step_text = "When region is ""North"": [line 143]" ' vla:2228
            Exit Function
        Case 81
            vla_step_text = "Set region-label to ""cold"". [line 144]" ' vla:2229
            Exit Function
        Case 82
            vla_step_text = "Set region-label to ""warm"". [line 147]" ' vla:2230
            Exit Function
        Case 83
            vla_step_text = "Set region-label to ""unknown"". [line 150]" ' vla:2231
            Exit Function
        Case 84
            vla_step_text = "Log ""verdict "" joined with verdict joined with "", region "" joined with region-label. [line 152]" ' vla:2232
            Exit Function
        Case 85
            vla_step_text = "Create a number called until-count. [line 154]" ' vla:2233
            Exit Function
        Case 86
            vla_step_text = "Set fuel to 3. [line 155]" ' vla:2234
            Exit Function
        Case 87
            vla_step_text = "Repeat until fuel is 0: [line 156]" ' vla:2235
            Exit Function
        Case 88
            vla_step_text = "Decrease fuel by 1. [line 157]" ' vla:2236
            Exit Function
        Case 89
            vla_step_text = "Increase until-count by 1. [line 158]" ' vla:2237
            Exit Function
        Case 90
            vla_step_text = "Log ""repeat-until ran "" joined with until-count joined with "" times"". [line 160]" ' vla:2238
            Exit Function
        Case 91
            vla_step_text = "Create a text called rescue. [line 166]" ' vla:2239
            Exit Function
        Case 92
            vla_step_text = "Try: [line 167]" ' vla:2240
            Exit Function
        Case 93
            vla_step_text = "Go to sheet Nowhere-Land. [line 168]" ' vla:2241
            Exit Function
        Case 94
            vla_step_text = "Set rescue to ""unreachable"". [line 169]" ' vla:2242
            Exit Function
        Case 95
            vla_step_text = "Log ""the problem was "" joined with the problem. [line 172]" ' vla:2243
            Exit Function
        Case 96
            vla_step_text = "If the problem is not empty, set rescue to ""rescued"". [line 173]" ' vla:2244
            Exit Function
        Case 97
            vla_step_text = "Create a number called risk-free. [line 175]" ' vla:2245
            Exit Function
        Case 98
            vla_step_text = "Try: [line 176]" ' vla:2246
            Exit Function
        Case 99
            vla_step_text = "Set risk-free to 7. [line 177]" ' vla:2247
            Exit Function
        Case 100
            vla_step_text = "Set risk-free to -1. [line 180]" ' vla:2248
            Exit Function
        Case 101
            vla_step_text = "Log ""rescue "" joined with rescue joined with "", risk-free "" joined with risk-free. [line 182]" ' vla:2249
            Exit Function
        Case 102
            vla_step_text = "Give back amount times 0.08. [line 187]" ' vla:2250
            Exit Function
        Case 103
            vla_step_text = "Set fee to tax of 100. [line 189]" ' vla:2251
            Exit Function
        Case 104
            vla_step_text = "Log ""fee is "" joined with fee. [line 190]" ' vla:2252
            Exit Function
        Case 105
            vla_step_text = "Create a text called fee-size. [line 191]" ' vla:2253
            Exit Function
        Case 106
            vla_step_text = "If tax of 50 is greater than 3: [line 192]" ' vla:2254
            Exit Function
        Case 107
            vla_step_text = "Set fee-size to ""big"". [line 193]" ' vla:2255
            Exit Function
        Case 108
            vla_step_text = "Set fee-size to ""small"". [line 196]" ' vla:2256
            Exit Function
        Case 109
            vla_step_text = "Fit column A. [line 199]" ' vla:2257
            Exit Function
        Case 110
            vla_step_text = "Fit column B. [line 200]" ' vla:2258
            Exit Function
        Case 111
            vla_step_text = "Set font-color of cell A1 to hot-pink. [line 201]" ' vla:2259
            Exit Function
        Case 112
            vla_step_text = "Give back sale times rate. [line 206]" ' vla:2260
            Exit Function
        Case 113
            vla_step_text = "Set full-commission to commission using sale of 2000 and rate of 10%. [line 208]" ' vla:2261
            Exit Function
        Case 114
            vla_step_text = "Set default-commission to get commission using sale of 600. [line 209]" ' vla:2262
            Exit Function
        Case 115
            vla_step_text = "Log ""commissions "" joined with full-commission joined with "" / "" joined with default-commission. [line 210]" ' vla:2263
            Exit Function
        Case 116
            vla_step_text = "Try: [line 219]" ' vla:2264
            Exit Function
        Case 117
            vla_step_text = "Add sheet called ""Q1 Data"". [line 220]" ' vla:2265
            Exit Function
        Case 118
            vla_step_text = "Go to sheet Output. [line 222]" ' vla:2266
            Exit Function
        Case 119
            vla_step_text = "Put ""spaced"" into cell 'Q1 Data'!A1. [line 223]" ' vla:2267
            Exit Function
        Case 120
            vla_step_text = "Create a text called spaced-check. [line 224]" ' vla:2268
            Exit Function
        Case 121
            vla_step_text = "Set spaced-check to value in cell 'Q1 Data'!A1. [line 225]" ' vla:2269
            Exit Function
        Case 122
            vla_step_text = "Give back 20%. [line 230]" ' vla:2270
            Exit Function
        Case 123
            vla_step_text = "Create a number called growth-check. [line 237]" ' vla:2271
            Exit Function
        Case 124
            vla_step_text = "Set growth-check to 200. [line 238]" ' vla:2272
            Exit Function
        Case 125
            vla_step_text = "Grow growth-check by 10%. [line 239]" ' vla:2273
            Exit Function
        Case 126
            vla_step_text = "Increase growth-check by 50%. [line 240]" ' vla:2274
            Exit Function
        Case 127
            vla_step_text = "Add 100 percent to growth-check. [line 241]" ' vla:2275
            Exit Function
        Case 128
            vla_step_text = "Decrease growth-check by 75%. [line 242]" ' vla:2276
            Exit Function
        Case 129
            vla_step_text = "Set pick-check to item 2 of found-items. [line 243]" ' vla:2277
            Exit Function
        Case 130
            vla_step_text = "Log ""grew to "" joined with growth-check joined with "", picked "" joined with pick-check joined with "", first "" joined with first of found-items. [line 244]" ' vla:2278
            Exit Function
        Case 131
            vla_step_text = "Create a text called quote-check. [line 245]" ' vla:2279
            Exit Function
        Case 132
            vla_step_text = "Set quote-check to ""He said """"ok"""""". [line 246]" ' vla:2280
            Exit Function
        Case 133
            vla_step_text = "Create a lookup called prices. [line 260]" ' vla:2281
            Exit Function
        Case 134
            vla_step_text = "Store 100 at key ""ax-7"" in prices. [line 261]" ' vla:2282
            Exit Function
        Case 135
            vla_step_text = "Store 250 under ""bx-2"" in prices. [line 262]" ' vla:2283
            Exit Function
        Case 136
            vla_step_text = "Store 120 at key ""AX-7"" in prices. [line 263]" ' vla:2284
            Exit Function
        Case 137
            vla_step_text = "Set ax-price to prices for ""ax-7"". [line 264]" ' vla:2285
            Exit Function
        Case 138
            vla_step_text = "Set key-count to count of keys of prices. [line 265]" ' vla:2286
            Exit Function
        Case 139
            vla_step_text = "Set price-sum to prices for ""ax-7"" plus prices for ""bx-2"". [line 266]" ' vla:2287
            Exit Function
        Case 140
            vla_step_text = "Create a text called key-list. [line 267]" ' vla:2288
            Exit Function
        Case 141
            vla_step_text = "For each k in keys of prices: [line 268]" ' vla:2289
            Exit Function
        Case 142
            vla_step_text = "Set key-list to key-list joined with k. [line 269]" ' vla:2290
            Exit Function
        Case 143
            vla_step_text = "Create a text called price-verdict. [line 271]" ' vla:2291
            Exit Function
        Case 144
            vla_step_text = "If prices for ""bx-2"" is greater than 200, set price-verdict to ""steep"". [line 272]" ' vla:2292
            Exit Function
        Case 145
            vla_step_text = "Create a text called pair-trace. [line 273]" ' vla:2293
            Exit Function
        Case 146
            vla_step_text = "For each pair in prices: [line 274]" ' vla:2294
            Exit Function
        Case 147
            vla_step_text = "If value of pair is greater than 200, set pair-trace to pair-trace joined with key of pair. [line 275]" ' vla:2295
            Exit Function
        Case 148
            vla_step_text = "Log ""lookup: ax "" joined with ax-price joined with "", keys "" joined with key-list joined with "", pairs "" joined with pair-trace. [line 277]" ' vla:2296
            Exit Function
        Case 149
            vla_step_text = "Put ""Item"" into cell J1. [line 284]" ' vla:2297
            Exit Function
        Case 150
            vla_step_text = "Put ""Amount"" into cell K1. [line 285]" ' vla:2298
            Exit Function
        Case 151
            vla_step_text = "Put ""Widget"" into cell J2. [line 286]" ' vla:2299
            Exit Function
        Case 152
            vla_step_text = "Put 10 into cell K2. [line 287]" ' vla:2300
            Exit Function
        Case 153
            vla_step_text = "Put ""Gadget"" into cell J3. [line 288]" ' vla:2301
            Exit Function
        Case 154
            vla_step_text = "Put 20 into cell K3. [line 289]" ' vla:2302
            Exit Function
        Case 155
            vla_step_text = "Turn J1:K3 into a table called SalesTable. [line 290]" ' vla:2303
            Exit Function
        Case 156
            vla_step_text = "Set style of table SalesTable to TableStyleMedium9. [line 291]" ' vla:2304
            Exit Function
        Case 157
            vla_step_text = "Show the total row of table SalesTable. [line 292]" ' vla:2305
            Exit Function
        Case 158
            vla_step_text = "Put ""X"" into cell J5. [line 296]" ' vla:2306
            Exit Function
        Case 159
            vla_step_text = "Put ""Y"" into cell J6. [line 297]" ' vla:2307
            Exit Function
        Case 160
            vla_step_text = "Turn J5:J6 into a table called QuietTable. [line 298]" ' vla:2308
            Exit Function
        Case 161
            vla_step_text = "Show the total row of table QuietTable. [line 299]" ' vla:2309
            Exit Function
        Case 162
            vla_step_text = "Hide the total row of table QuietTable. [line 300]" ' vla:2310
            Exit Function
        Case 163
            vla_step_text = "Put ""A"" into cell J8. [line 304]" ' vla:2311
            Exit Function
        Case 164
            vla_step_text = "Put ""B"" into cell J9. [line 305]" ' vla:2312
            Exit Function
        Case 165
            vla_step_text = "Turn J8:J9 into a table called TempTable. [line 306]" ' vla:2313
            Exit Function
        Case 166
            vla_step_text = "Turn table TempTable back into a range. [line 307]" ' vla:2314
            Exit Function
        Case 167
            vla_step_text = "Put ""Item"" into cell J11. [line 312]" ' vla:2315
            Exit Function
        Case 168
            vla_step_text = "Put ""Qty"" into cell K11. [line 313]" ' vla:2316
            Exit Function
        Case 169
            vla_step_text = "Put ""Bolt"" into cell J12. [line 314]" ' vla:2317
            Exit Function
        Case 170
            vla_step_text = "Put 5 into cell K12. [line 315]" ' vla:2318
            Exit Function
        Case 171
            vla_step_text = "Put ""Nut"" into cell J13. [line 316]" ' vla:2319
            Exit Function
        Case 172
            vla_step_text = "Put 8 into cell K13. [line 317]" ' vla:2320
            Exit Function
        Case 173
            vla_step_text = "Turn J11:K13 into a table called EditTable. [line 318]" ' vla:2321
            Exit Function
        Case 174
            vla_step_text = "Add a row to table EditTable. [line 319]" ' vla:2322
            Exit Function
        Case 175
            vla_step_text = "Delete row 1 of table EditTable. [line 320]" ' vla:2323
            Exit Function
        Case 176
            vla_step_text = "Set qty-values to column Qty of table EditTable. [line 321]" ' vla:2324
            Exit Function
        Case 177
            vla_step_text = "For each q in qty-values, log q. [line 322]" ' vla:2325
            Exit Function
        Case 178
            vla_step_text = "Put ""Region"" into cell N1. [line 332]" ' vla:2326
            Exit Function
        Case 179
            vla_step_text = "Put ""Product"" into cell O1. [line 333]" ' vla:2327
            Exit Function
        Case 180
            vla_step_text = "Put ""Segment"" into cell P1. [line 334]" ' vla:2328
            Exit Function
        Case 181
            vla_step_text = "Put ""Channel"" into cell Q1. [line 335]" ' vla:2329
            Exit Function
        Case 182
            vla_step_text = "Put ""Units"" into cell R1. [line 336]" ' vla:2330
            Exit Function
        Case 183
            vla_step_text = "Put ""Revenue"" into cell S1. [line 337]" ' vla:2331
            Exit Function
        Case 184
            vla_step_text = "Put ""North"" into cell N2. [line 338]" ' vla:2332
            Exit Function
        Case 185
            vla_step_text = "Put ""Widget"" into cell O2. [line 339]" ' vla:2333
            Exit Function
        Case 186
            vla_step_text = "Put ""Retail"" into cell P2. [line 340]" ' vla:2334
            Exit Function
        Case 187
            vla_step_text = "Put ""Online"" into cell Q2. [line 341]" ' vla:2335
            Exit Function
        Case 188
            vla_step_text = "Put 10 into cell R2. [line 342]" ' vla:2336
            Exit Function
        Case 189
            vla_step_text = "Put 500 into cell S2. [line 343]" ' vla:2337
            Exit Function
        Case 190
            vla_step_text = "Put ""North"" into cell N3. [line 344]" ' vla:2338
            Exit Function
        Case 191
            vla_step_text = "Put ""Gadget"" into cell O3. [line 345]" ' vla:2339
            Exit Function
        Case 192
            vla_step_text = "Put ""Wholesale"" into cell P3. [line 346]" ' vla:2340
            Exit Function
        Case 193
            vla_step_text = "Put ""Store"" into cell Q3. [line 347]" ' vla:2341
            Exit Function
        Case 194
            vla_step_text = "Put 5 into cell R3. [line 348]" ' vla:2342
            Exit Function
        Case 195
            vla_step_text = "Put 200 into cell S3. [line 349]" ' vla:2343
            Exit Function
        Case 196
            vla_step_text = "Put ""South"" into cell N4. [line 350]" ' vla:2344
            Exit Function
        Case 197
            vla_step_text = "Put ""Widget"" into cell O4. [line 351]" ' vla:2345
            Exit Function
        Case 198
            vla_step_text = "Put ""Wholesale"" into cell P4. [line 352]" ' vla:2346
            Exit Function
        Case 199
            vla_step_text = "Put ""Online"" into cell Q4. [line 353]" ' vla:2347
            Exit Function
        Case 200
            vla_step_text = "Put 20 into cell R4. [line 354]" ' vla:2348
            Exit Function
        Case 201
            vla_step_text = "Put 900 into cell S4. [line 355]" ' vla:2349
            Exit Function
        Case 202
            vla_step_text = "Put ""South"" into cell N5. [line 356]" ' vla:2350
            Exit Function
        Case 203
            vla_step_text = "Put ""Gadget"" into cell O5. [line 357]" ' vla:2351
            Exit Function
        Case 204
            vla_step_text = "Put ""Retail"" into cell P5. [line 358]" ' vla:2352
            Exit Function
        Case 205
            vla_step_text = "Put ""Store"" into cell Q5. [line 359]" ' vla:2353
            Exit Function
        Case 206
            vla_step_text = "Put 8 into cell R5. [line 360]" ' vla:2354
            Exit Function
        Case 207
            vla_step_text = "Put 300 into cell S5. [line 361]" ' vla:2355
            Exit Function
        Case 208
            vla_step_text = "Put ""East"" into cell N6. [line 362]" ' vla:2356
            Exit Function
        Case 209
            vla_step_text = "Put ""Widget"" into cell O6. [line 363]" ' vla:2357
            Exit Function
        Case 210
            vla_step_text = "Put ""Retail"" into cell P6. [line 364]" ' vla:2358
            Exit Function
        Case 211
            vla_step_text = "Put ""Online"" into cell Q6. [line 365]" ' vla:2359
            Exit Function
        Case 212
            vla_step_text = "Put 12 into cell R6. [line 366]" ' vla:2360
            Exit Function
        Case 213
            vla_step_text = "Put 600 into cell S6. [line 367]" ' vla:2361
            Exit Function
        Case 214
            vla_step_text = "Put ""East"" into cell N7. [line 368]" ' vla:2362
            Exit Function
        Case 215
            vla_step_text = "Put ""Gadget"" into cell O7. [line 369]" ' vla:2363
            Exit Function
        Case 216
            vla_step_text = "Put ""Wholesale"" into cell P7. [line 370]" ' vla:2364
            Exit Function
        Case 217
            vla_step_text = "Put ""Store"" into cell Q7. [line 371]" ' vla:2365
            Exit Function
        Case 218
            vla_step_text = "Put 6 into cell R7. [line 372]" ' vla:2366
            Exit Function
        Case 219
            vla_step_text = "Put 250 into cell S7. [line 373]" ' vla:2367
            Exit Function
        Case 220
            vla_step_text = "Make a pivot table from N1:S7 at U1 called SalesPivot. [line 381]" ' vla:2368
            Exit Function
        Case 221
            vla_step_text = "Add rows of Region, Product to pivot SalesPivot. [line 382]" ' vla:2369
            Exit Function
        Case 222
            vla_step_text = "Add columns of Segment to pivot SalesPivot. [line 383]" ' vla:2370
            Exit Function
        Case 223
            vla_step_text = "Add filters of Channel to pivot SalesPivot. [line 384]" ' vla:2371
            Exit Function
        Case 224
            vla_step_text = "Add Revenue to pivot SalesPivot as a sum. [line 392]" ' vla:2372
            Exit Function
        Case 225
            vla_step_text = "Add Revenue, Units to pivot SalesPivot as a count. [line 393]" ' vla:2373
            Exit Function
        Case 226
            vla_step_text = "Add Units to pivot SalesPivot as an average. [line 394]" ' vla:2374
            Exit Function
        Case 227
            vla_step_text = "Make a pivot table from N1:S7 at N20 called FullPivot with rows of Region, Product, and Channel and columns of Segment and values of Revenue, Units. [line 406]" ' vla:2375
            Exit Function
        Case 228
            vla_step_text = "Refresh pivot SalesPivot. [line 415]" ' vla:2376
            Exit Function
        Case 229
            vla_step_text = "Refresh every pivot table. [line 416]" ' vla:2377
            Exit Function
        Case 230
            vla_step_text = "Collapse Region in pivot SalesPivot. [line 417]" ' vla:2378
            Exit Function
        Case 231
            vla_step_text = "Collapse Product in pivot FullPivot. [line 427]" ' vla:2379
            Exit Function
        Case 232
            vla_step_text = "Expand Product in pivot FullPivot. [line 428]" ' vla:2380
            Exit Function
        Case 233
            vla_step_text = "Show pivot SalesPivot in tabular form. [line 443]" ' vla:2381
            Exit Function
        Case 234
            vla_step_text = "Show pivot SalesPivot in compact form. [line 444]" ' vla:2382
            Exit Function
        Case 235
            vla_step_text = "Show pivot FullPivot in outline form. [line 445]" ' vla:2383
            Exit Function
        Case 236
            vla_step_text = "Hide subtotals for Region in pivot SalesPivot. [line 453]" ' vla:2384
            Exit Function
        Case 237
            vla_step_text = "Hide subtotals for Product in pivot FullPivot. [line 454]" ' vla:2385
            Exit Function
        Case 238
            vla_step_text = "Show subtotals for Product in pivot FullPivot. [line 455]" ' vla:2386
            Exit Function
        Case 239
            vla_step_text = "Add a blank row after Region in pivot SalesPivot. [line 457]" ' vla:2387
            Exit Function
        Case 240
            vla_step_text = "Add a blank row after Product in pivot FullPivot. [line 458]" ' vla:2388
            Exit Function
        Case 241
            vla_step_text = "Remove a blank row after Product in pivot FullPivot. [line 459]" ' vla:2389
            Exit Function
        Case 242
            vla_step_text = "Sort Region in pivot SalesPivot descending. [line 476]" ' vla:2390
            Exit Function
        Case 243
            vla_step_text = "Sort Region in pivot SalesPivot ascending. [line 477]" ' vla:2391
            Exit Function
        Case 244
            vla_step_text = "Sort Product in pivot FullPivot descending by Revenue. [line 478]" ' vla:2392
            Exit Function
        Case 245
            vla_step_text = "Sort Product in pivot FullPivot ascending by Revenue. [line 479]" ' vla:2393
            Exit Function
        Case 246
            vla_step_text = "Make a pivot table from N1:S7 at N200 called RenameMePivot. [line 490]" ' vla:2394
            Exit Function
        Case 247
            vla_step_text = "Rename pivot RenameMePivot to RenamedPivot. [line 491]" ' vla:2395
            Exit Function
        Case 248
            vla_step_text = "Make a pivot table from N1:S7 at N220 called ClearMePivot. [line 493]" ' vla:2396
            Exit Function
        Case 249
            vla_step_text = "Add rows of Region to pivot ClearMePivot. [line 494]" ' vla:2397
            Exit Function
        Case 250
            vla_step_text = "Add Revenue to pivot ClearMePivot as a sum. [line 495]" ' vla:2398
            Exit Function
        Case 251
            vla_step_text = "Clear pivot ClearMePivot. [line 496]" ' vla:2399
            Exit Function
        Case 252
            vla_step_text = "Make a pivot table from N1:S7 at N240 called RemoveFieldMePivot. [line 498]" ' vla:2400
            Exit Function
        Case 253
            vla_step_text = "Add rows of Region, Product to pivot RemoveFieldMePivot. [line 499]" ' vla:2401
            Exit Function
        Case 254
            vla_step_text = "Remove Product from pivot RemoveFieldMePivot. [line 500]" ' vla:2402
            Exit Function
        Case 255
            vla_step_text = "Put ""West"" into cell N8. [line 506]" ' vla:2403
            Exit Function
        Case 256
            vla_step_text = "Put ""Widget"" into cell O8. [line 507]" ' vla:2404
            Exit Function
        Case 257
            vla_step_text = "Put ""Retail"" into cell P8. [line 508]" ' vla:2405
            Exit Function
        Case 258
            vla_step_text = "Put ""Online"" into cell Q8. [line 509]" ' vla:2406
            Exit Function
        Case 259
            vla_step_text = "Put 15 into cell R8. [line 510]" ' vla:2407
            Exit Function
        Case 260
            vla_step_text = "Put 700 into cell S8. [line 511]" ' vla:2408
            Exit Function
        Case 261
            vla_step_text = "Make a pivot table from N1:S7 at N260 called SourceTestPivot. [line 513]" ' vla:2409
            Exit Function
        Case 262
            vla_step_text = "Add rows of Region to pivot SourceTestPivot. [line 514]" ' vla:2410
            Exit Function
        Case 263
            vla_step_text = "Change the source of pivot SourceTestPivot to N1:S8. [line 515]" ' vla:2411
            Exit Function
        Case 264
            vla_step_text = "Make a pivot table from N1:S7 at N40 called TempPivot. [line 520]" ' vla:2412
            Exit Function
        Case 265
            vla_step_text = "Delete pivot TempPivot. [line 521]" ' vla:2413
            Exit Function
        Case 266
            vla_step_text = "Put grand into cell H1. [line 524]" ' vla:2414
            Exit Function
        Case 267
            vla_step_text = "Put biggest into cell H2. [line 525]" ' vla:2415
            Exit Function
        Case 268
            vla_step_text = "Put round-check into cell H3. [line 526]" ' vla:2416
            Exit Function
        Case 269
            vla_step_text = "Put thousand-check into cell H4. [line 527]" ' vla:2417
            Exit Function
        Case 270
            vla_step_text = "Put search-row into cell H5. [line 528]" ' vla:2418
            Exit Function
        Case 271
            vla_step_text = "Put f-last into cell H6. [line 529]" ' vla:2419
            Exit Function
        Case 272
            vla_step_text = "Put list-count into cell H7. [line 530]" ' vla:2420
            Exit Function
        Case 273
            vla_step_text = "Put verdict into cell H8. [line 531]" ' vla:2421
            Exit Function
        Case 274
            vla_step_text = "Put region-label into cell H9. [line 532]" ' vla:2422
            Exit Function
        Case 275
            vla_step_text = "Put until-count into cell H10. [line 533]" ' vla:2423
            Exit Function
        Case 276
            vla_step_text = "Put rescue into cell H11. [line 534]" ' vla:2424
            Exit Function
        Case 277
            vla_step_text = "Put risk-free into cell H12. [line 535]" ' vla:2425
            Exit Function
        Case 278
            vla_step_text = "Put fee into cell H13. [line 536]" ' vla:2426
            Exit Function
        Case 279
            vla_step_text = "Put fee-size into cell H14. [line 537]" ' vla:2427
            Exit Function
        Case 280
            vla_step_text = "Put num-col-check into cell H15. [line 538]" ' vla:2428
            Exit Function
        Case 281
            vla_step_text = "Put value-word-check into cell H17. [line 539]" ' vla:2429
            Exit Function
        Case 282
            vla_step_text = "Put spaced-check into cell H18. [line 540]" ' vla:2430
            Exit Function
        Case 283
            vla_step_text = "Put full-commission into cell H19. [line 541]" ' vla:2431
            Exit Function
        Case 284
            vla_step_text = "Put default-commission into cell H20. [line 542]" ' vla:2432
            Exit Function
        Case 285
            vla_step_text = "Put growth-check into cell H21. [line 543]" ' vla:2433
            Exit Function
        Case 286
            vla_step_text = "Put pick-check into cell H22. [line 544]" ' vla:2434
            Exit Function
        Case 287
            vla_step_text = "Put quote-check into cell H23. [line 545]" ' vla:2435
            Exit Function
        Case 288
            vla_step_text = "Put vat-rate into cell H24. [line 546]" ' vla:2436
            Exit Function
        Case 289
            vla_step_text = "Put ax-price into cell H26. [line 547]" ' vla:2437
            Exit Function
        Case 290
            vla_step_text = "Put key-count into cell H27. [line 548]" ' vla:2438
            Exit Function
        Case 291
            vla_step_text = "Put key-list into cell H28. [line 549]" ' vla:2439
            Exit Function
        Case 292
            vla_step_text = "Put price-sum into cell H29. [line 550]" ' vla:2440
            Exit Function
        Case 293
            vla_step_text = "Put price-verdict into cell H30. [line 551]" ' vla:2441
            Exit Function
        Case 294
            vla_step_text = "Put pair-trace into cell H31. [line 552]" ' vla:2442
            Exit Function
        Case 295
            vla_step_text = "(set! (range ""h25"") ""vla-row"") [line 559]" ' vla:2443
            Exit Function
        Case 296
            vla_step_text = "Repeat 2 times: [line 560]" ' vla:2444
            Exit Function
        Case 297
            vla_step_text = "(debug-print counter) [line 561]" ' vla:2445
            Exit Function
        Case 298
            vla_step_text = "Work on sheet Demo. [line 576]" ' vla:2446
            Exit Function
        Case 299
            vla_step_text = "Make cell A1 italic. [line 579]" ' vla:2447
            Exit Function
        Case 300
            vla_step_text = "Make cell A2 red. [line 580]" ' vla:2448
            Exit Function
        Case 301
            vla_step_text = "Make cell A3 yellow. [line 581]" ' vla:2449
            Exit Function
        Case 302
            vla_step_text = "Set font-color of cell A4 to hot-pink. [line 582]" ' vla:2450
            Exit Function
        Case 303
            vla_step_text = "Set fill-color of cell A5 to ""#FF69B4"". [line 583]" ' vla:2451
            Exit Function
        Case 304
            vla_step_text = "Clear color of cell A5. [line 584]" ' vla:2452
            Exit Function
        Case 305
            vla_step_text = "Add border to range A1:E10. [line 585]" ' vla:2453
            Exit Function
        Case 306
            vla_step_text = "Format cell B1 as currency. [line 586]" ' vla:2454
            Exit Function
        Case 307
            vla_step_text = "Format cell B2 as percent. [line 587]" ' vla:2455
            Exit Function
        Case 308
            vla_step_text = "Format cell B3 as date. [line 588]" ' vla:2456
            Exit Function
        Case 309
            vla_step_text = "Center cell A1. [line 589]" ' vla:2457
            Exit Function
        Case 310
            vla_step_text = "Center range A1:E1. [line 590]" ' vla:2458
            Exit Function
        Case 311
            vla_step_text = "Align range A2:A5 left. [line 591]" ' vla:2459
            Exit Function
        Case 312
            vla_step_text = "Align cell A6 center. [line 592]" ' vla:2460
            Exit Function
        Case 313
            vla_step_text = "Wrap text in range C1:C5. [line 593]" ' vla:2461
            Exit Function
        Case 314
            vla_step_text = "Unwrap text in range C1:C5. [line 594]" ' vla:2462
            Exit Function
        Case 315
            vla_step_text = "Merge range D1:D3. [line 595]" ' vla:2463
            Exit Function
        Case 316
            vla_step_text = "Unmerge range D1:D3. [line 596]" ' vla:2464
            Exit Function
        Case 317
            vla_step_text = "Set height of row 1 to 30. [line 601]" ' vla:2465
            Exit Function
        Case 318
            vla_step_text = "Set width of column A to 20. [line 602]" ' vla:2466
            Exit Function
        Case 319
            vla_step_text = "Insert column before B. [line 603]" ' vla:2467
            Exit Function
        Case 320
            vla_step_text = "Hide column C. [line 604]" ' vla:2468
            Exit Function
        Case 321
            vla_step_text = "Unhide column C. [line 605]" ' vla:2469
            Exit Function
        Case 322
            vla_step_text = "Delete column B. [line 606]" ' vla:2470
            Exit Function
        Case 323
            vla_step_text = "Insert row at 5. [line 607]" ' vla:2471
            Exit Function
        Case 324
            vla_step_text = "Insert a row. [line 608]" ' vla:2472
            Exit Function
        Case 325
            vla_step_text = "Delete row 2. [line 609]" ' vla:2473
            Exit Function
        Case 326
            vla_step_text = "Delete the third row. [line 610]" ' vla:2474
            Exit Function
        Case 327
            vla_step_text = "Hide row 10. [line 611]" ' vla:2475
            Exit Function
        Case 328
            vla_step_text = "Unhide row 10. [line 612]" ' vla:2476
            Exit Function
        Case 329
            vla_step_text = "Freeze top row. [line 613]" ' vla:2477
            Exit Function
        Case 330
            vla_step_text = "Unfreeze panes. [line 614]" ' vla:2478
            Exit Function
        Case 331
            vla_step_text = "Fit column A. [line 615]" ' vla:2479
            Exit Function
        Case 332
            vla_step_text = "Fit all columns. [line 616]" ' vla:2480
            Exit Function
        Case 333
            vla_step_text = "Put ""Region"" into cell G1. [line 620]" ' vla:2481
            Exit Function
        Case 334
            vla_step_text = "Put ""Amount"" into cell H1. [line 621]" ' vla:2482
            Exit Function
        Case 335
            vla_step_text = "Put ""Notes"" into cell I1. [line 622]" ' vla:2483
            Exit Function
        Case 336
            vla_step_text = "Put ""West"" into cell G2. [line 623]" ' vla:2484
            Exit Function
        Case 337
            vla_step_text = "Put 100 into cell H2. [line 624]" ' vla:2485
            Exit Function
        Case 338
            vla_step_text = "Put ""ok"" into cell I2. [line 625]" ' vla:2486
            Exit Function
        Case 339
            vla_step_text = "Put ""East"" into cell G3. [line 626]" ' vla:2487
            Exit Function
        Case 340
            vla_step_text = "Put 250 into cell H3. [line 627]" ' vla:2488
            Exit Function
        Case 341
            vla_step_text = "Put ""ok"" into cell I3. [line 628]" ' vla:2489
            Exit Function
        Case 342
            vla_step_text = "Put ""West"" into cell G4. [line 629]" ' vla:2490
            Exit Function
        Case 343
            vla_step_text = "Put 100 into cell H4. [line 630]" ' vla:2491
            Exit Function
        Case 344
            vla_step_text = "Put ""dup"" into cell I4. [line 631]" ' vla:2492
            Exit Function
        Case 345
            vla_step_text = "Sort range G1:I4 by column H1 descending. [line 632]" ' vla:2493
            Exit Function
        Case 346
            vla_step_text = "Keep only rows of range G1:I4 where column 1 is ""West"". [line 633]" ' vla:2494
            Exit Function
        Case 347
            vla_step_text = "Show all rows. [line 634]" ' vla:2495
            Exit Function
        Case 348
            vla_step_text = "Replace ""dup"" with ""ok"" in range G1:I4. [line 635]" ' vla:2496
            Exit Function
        Case 349
            vla_step_text = "Remove duplicates from range G1:I4 by column 1. [line 636]" ' vla:2497
            Exit Function
        Case 350
            vla_step_text = "Replace ""ok"" with ""fine"" in column I. [line 637]" ' vla:2498
            Exit Function
        Case 351
            vla_step_text = "Convert range G1:I4 to values. [line 638]" ' vla:2499
            Exit Function
        Case 352
            vla_step_text = "Clear formatting of range G1:I4. [line 639]" ' vla:2500
            Exit Function
        Case 353
            vla_step_text = "Name range G1:I4 as demo_table. [line 640]" ' vla:2501
            Exit Function
        Case 354
            vla_step_text = "Copy range G1:I4 to range K1:M4. [line 641]" ' vla:2502
            Exit Function
        Case 355
            vla_step_text = "Set tab-color of sheet Demo to hot-pink. [line 645]" ' vla:2503
            Exit Function
        Case 356
            vla_step_text = "Protect this sheet with password ""demo123"". [line 646]" ' vla:2504
            Exit Function
        Case 357
            vla_step_text = "Unprotect this sheet with password ""demo123"". [line 647]" ' vla:2505
            Exit Function
        Case 358
            vla_step_text = "Try: [line 675]" ' vla:2506
            Exit Function
        Case 359
            vla_step_text = "Delete sheet GStruct. [line 676]" ' vla:2507
            Exit Function
        Case 360
            vla_step_text = "Work on sheet GStruct. [line 678]" ' vla:2508
            Exit Function
        Case 361
            vla_step_text = "Hide row 3. [line 680]" ' vla:2509
            Exit Function
        Case 362
            vla_step_text = "Hide column B. [line 681]" ' vla:2510
            Exit Function
        Case 363
            vla_step_text = "Unhide all rows and columns. [line 682]" ' vla:2511
            Exit Function
        Case 364
            vla_step_text = "Set font size of cell A6 to 36. [line 684]" ' vla:2512
            Exit Function
        Case 365
            vla_step_text = "Fit row 6. [line 685]" ' vla:2513
            Exit Function
        Case 366
            vla_step_text = "Group rows 10 through 12. [line 687]" ' vla:2514
            Exit Function
        Case 367
            vla_step_text = "Group rows 14 through 16. [line 688]" ' vla:2515
            Exit Function
        Case 368
            vla_step_text = "Ungroup rows 14 through 16. [line 689]" ' vla:2516
            Exit Function
        Case 369
            vla_step_text = "Put ""before-insert"" into cell A20. [line 691]" ' vla:2517
            Exit Function
        Case 370
            vla_step_text = "Insert 3 rows at row 20. [line 692]" ' vla:2518
            Exit Function
        Case 371
            vla_step_text = "Put ""before-delete"" into cell A40. [line 694]" ' vla:2519
            Exit Function
        Case 372
            vla_step_text = "Delete rows 38 through 39. [line 695]" ' vla:2520
            Exit Function
        Case 373
            vla_step_text = "Put ""marker-b"" into cell B50. [line 697]" ' vla:2521
            Exit Function
        Case 374
            vla_step_text = "Put ""marker-c"" into cell C50. [line 698]" ' vla:2522
            Exit Function
        Case 375
            vla_step_text = "Put ""marker-d"" into cell D50. [line 699]" ' vla:2523
            Exit Function
        Case 376
            vla_step_text = "Move column B before column D. [line 700]" ' vla:2524
            Exit Function
        Case 377
            vla_step_text = "Freeze the first 2 rows. [line 702]" ' vla:2525
            Exit Function
        Case 378
            vla_step_text = "Make cell A60 bold. [line 705]" ' vla:2526
            Exit Function
        Case 379
            vla_step_text = "Make cell A60 red. [line 706]" ' vla:2527
            Exit Function
        Case 380
            vla_step_text = "Copy formatting of A60:A60 to C60:C60. [line 707]" ' vla:2528
            Exit Function
        Case 381
            vla_step_text = "Make C62:C62 look like A60:A60. [line 708]" ' vla:2529
            Exit Function
        Case 382
            vla_step_text = "Put formula ""=5*2"" into cell B64. [line 710]" ' vla:2530
            Exit Function
        Case 383
            vla_step_text = "Copy formulas of B64 to D64. [line 711]" ' vla:2531
            Exit Function
        Case 384
            vla_step_text = "Set width of column F to 33. [line 713]" ' vla:2532
            Exit Function
        Case 385
            vla_step_text = "Copy column widths of F1:F1 to H1:H1. [line 714]" ' vla:2533
            Exit Function
        Case 386
            vla_step_text = "Put 1 into cell A68. [line 716]" ' vla:2534
            Exit Function
        Case 387
            vla_step_text = "Put 2 into cell A69. [line 717]" ' vla:2535
            Exit Function
        Case 388
            vla_step_text = "Put 3 into cell A70. [line 718]" ' vla:2536
            Exit Function
        Case 389
            vla_step_text = "Copy range A68:A70 to C68 transposed. [line 719]" ' vla:2537
            Exit Function
        Case 390
            vla_step_text = "Put ""cutme"" into cell A72. [line 721]" ' vla:2538
            Exit Function
        Case 391
            vla_step_text = "Move range A72:A72 to C72. [line 722]" ' vla:2539
            Exit Function
        Case 392
            vla_step_text = "Put ""rowdata"" into cell A74. [line 724]" ' vla:2540
            Exit Function
        Case 393
            vla_step_text = "Copy row 74 to row 76. [line 725]" ' vla:2541
            Exit Function
        Case 394
            vla_step_text = "Put ""clearme"" into cell A80. [line 728]" ' vla:2542
            Exit Function
        Case 395
            vla_step_text = "Make cell A80 bold. [line 729]" ' vla:2543
            Exit Function
        Case 396
            vla_step_text = "Make cell A80 red. [line 730]" ' vla:2544
            Exit Function
        Case 397
            vla_step_text = "Clear everything from A80:B80. [line 731]" ' vla:2545
            Exit Function
        Case 398
            vla_step_text = "Put ""m1"" into cell A84. [line 733]" ' vla:2546
            Exit Function
        Case 399
            vla_step_text = "Put ""m2"" into cell A85. [line 734]" ' vla:2547
            Exit Function
        Case 400
            vla_step_text = "Put ""m3"" into cell A86. [line 735]" ' vla:2548
            Exit Function
        Case 401
            vla_step_text = "Put ""m4"" into cell A87. [line 736]" ' vla:2549
            Exit Function
        Case 402
            vla_step_text = "Put ""m5"" into cell A88. [line 737]" ' vla:2550
            Exit Function
        Case 403
            vla_step_text = "Delete A84:A86 and shift cells up. [line 738]" ' vla:2551
            Exit Function
        Case 404
            vla_step_text = "Put ""x1"" into cell B90. [line 740]" ' vla:2552
            Exit Function
        Case 405
            vla_step_text = "Put ""x2"" into cell C90. [line 741]" ' vla:2553
            Exit Function
        Case 406
            vla_step_text = "Put ""x3"" into cell D90. [line 742]" ' vla:2554
            Exit Function
        Case 407
            vla_step_text = "Put ""rightdata"" into cell E90. [line 743]" ' vla:2555
            Exit Function
        Case 408
            vla_step_text = "Delete B90:D90 and shift cells left. [line 744]" ' vla:2556
            Exit Function
        Case 409
            vla_step_text = "Put 1 into cell A94. [line 746]" ' vla:2557
            Exit Function
        Case 410
            vla_step_text = "Put 2 into cell B94. [line 747]" ' vla:2558
            Exit Function
        Case 411
            vla_step_text = "Put 3 into cell A95. [line 748]" ' vla:2559
            Exit Function
        Case 412
            vla_step_text = "Put 5 into cell A96. [line 749]" ' vla:2560
            Exit Function
        Case 413
            vla_step_text = "Put 6 into cell B96. [line 750]" ' vla:2561
            Exit Function
        Case 414
            vla_step_text = "Delete blank rows in A94:B96. [line 751]" ' vla:2562
            Exit Function
        Case 415
            vla_step_text = "Put 1 into cell A150. [line 754]" ' vla:2563
            Exit Function
        Case 416
            vla_step_text = "Put 2 into cell A151. [line 755]" ' vla:2564
            Exit Function
        Case 417
            vla_step_text = "Put 3 into cell A152. [line 756]" ' vla:2565
            Exit Function
        Case 418
            vla_step_text = "Put 4 into cell A153. [line 757]" ' vla:2566
            Exit Function
        Case 419
            vla_step_text = "Band every other row of A150:D153 ""#D9D9D9"". [line 758]" ' vla:2567
            Exit Function
        Case 420
            vla_step_text = "Make row 165 a header row. [line 759]" ' vla:2568
            Exit Function
        Case 421
            vla_step_text = "Fill A170:A179 with a series starting at 1. [line 762]" ' vla:2569
            Exit Function
        Case 422
            vla_step_text = "Fill A180:A184 with a series starting at 1 with step 2. [line 763]" ' vla:2570
            Exit Function
        Case 423
            vla_step_text = "Fill A190:A194 with a growth series starting at 2. [line 764]" ' vla:2571
            Exit Function
        Case 424
            vla_step_text = "Fill A200:A203 with a growth series starting at 2 with step 3. [line 765]" ' vla:2572
            Exit Function
        Case 425
            vla_step_text = "Make cell Z500 bold. [line 777]" ' vla:2573
            Exit Function
        Case 426
            vla_step_text = "Clear everything from Z500. [line 778]" ' vla:2574
            Exit Function
        Case 427
            vla_step_text = "Remove trailing empty rows and columns. [line 779]" ' vla:2575
            Exit Function
        Case 428
            vla_step_text = "Try: [line 792]" ' vla:2576
            Exit Function
        Case 429
            vla_step_text = "Delete sheet GFormat. [line 793]" ' vla:2577
            Exit Function
        Case 430
            vla_step_text = "Work on sheet GFormat. [line 795]" ' vla:2578
            Exit Function
        Case 431
            vla_step_text = "Make range A1:C1 bold. [line 799]" ' vla:2579
            Exit Function
        Case 432
            vla_step_text = "Make range A2:C2 italic. [line 800]" ' vla:2580
            Exit Function
        Case 433
            vla_step_text = "Make range A3:C3 blue. [line 801]" ' vla:2581
            Exit Function
        Case 434
            vla_step_text = "Set font-color of range A4:C4 to ""#FF0000"". [line 802]" ' vla:2582
            Exit Function
        Case 435
            vla_step_text = "Set fill-color of range A5:C5 to ""#00FF00"". [line 803]" ' vla:2583
            Exit Function
        Case 436
            vla_step_text = "Make range A6:C6 yellow. [line 804]" ' vla:2584
            Exit Function
        Case 437
            vla_step_text = "Clear fill-color of range B6:C6. [line 805]" ' vla:2585
            Exit Function
        Case 438
            vla_step_text = "Set font size of range A7:C7 to 16. [line 806]" ' vla:2586
            Exit Function
        Case 439
            vla_step_text = "Make range A8:C8 bold. [line 809]" ' vla:2587
            Exit Function
        Case 440
            vla_step_text = "Make range B8:C8 not bold. [line 810]" ' vla:2588
            Exit Function
        Case 441
            vla_step_text = "Make range A9:C9 italic. [line 811]" ' vla:2589
            Exit Function
        Case 442
            vla_step_text = "Make cell C9 not italic. [line 812]" ' vla:2590
            Exit Function
        Case 443
            vla_step_text = "Underline range A10:C10. [line 813]" ' vla:2591
            Exit Function
        Case 444
            vla_step_text = "Strike through range A11:C11. [line 814]" ' vla:2592
            Exit Function
        Case 445
            vla_step_text = "Underline range A12:C12. [line 815]" ' vla:2593
            Exit Function
        Case 446
            vla_step_text = "Remove underline from range B12:C12. [line 816]" ' vla:2594
            Exit Function
        Case 447
            vla_step_text = "Strike through range A13:C13. [line 817]" ' vla:2595
            Exit Function
        Case 448
            vla_step_text = "Remove the strikethrough from cell C13. [line 818]" ' vla:2596
            Exit Function
        Case 449
            vla_step_text = "Set font of range A14:C14 to ""Courier New"". [line 819]" ' vla:2597
            Exit Function
        Case 450
            vla_step_text = "Align range A15:C15 to the top. [line 823]" ' vla:2598
            Exit Function
        Case 451
            vla_step_text = "Align range A16:C16 to the middle. [line 824]" ' vla:2599
            Exit Function
        Case 452
            vla_step_text = "Align range A17:C17 to the top. [line 825]" ' vla:2600
            Exit Function
        Case 453
            vla_step_text = "Align range B17:C17 to the bottom. [line 826]" ' vla:2601
            Exit Function
        Case 454
            vla_step_text = "Indent range A18:C18 by 2. [line 827]" ' vla:2602
            Exit Function
        Case 455
            vla_step_text = "Rotate text in range A19:C19 by 45 degrees. [line 828]" ' vla:2603
            Exit Function
        Case 456
            vla_step_text = "Add a border around range B21:D23. [line 834]" ' vla:2604
            Exit Function
        Case 457
            vla_step_text = "Add a bottom border to range B25:D25. [line 835]" ' vla:2605
            Exit Function
        Case 458
            vla_step_text = "Add borders to every cell in range B27:D29. [line 836]" ' vla:2606
            Exit Function
        Case 459
            vla_step_text = "Add borders to every cell in range B31:D33. [line 837]" ' vla:2607
            Exit Function
        Case 460
            vla_step_text = "Remove borders from range C31:D33. [line 838]" ' vla:2608
            Exit Function
        Case 461
            vla_step_text = "Set width of column F to 40. [line 846]" ' vla:2609
            Exit Function
        Case 462
            vla_step_text = "Put 1234.56 into cell F40. [line 847]" ' vla:2610
            Exit Function
        Case 463
            vla_step_text = "Format cell F40 as a number. [line 848]" ' vla:2611
            Exit Function
        Case 464
            vla_step_text = "Put 1234.56 into cell F41. [line 849]" ' vla:2612
            Exit Function
        Case 465
            vla_step_text = "Format cell F41 as a number with 3 decimals. [line 850]" ' vla:2613
            Exit Function
        Case 466
            vla_step_text = "Put 1234.56 into cell F42. [line 851]" ' vla:2614
            Exit Function
        Case 467
            vla_step_text = "Format cell F42 as a number with thousands separators. [line 852]" ' vla:2615
            Exit Function
        Case 468
            vla_step_text = "Put 1234.56 into cell F43. [line 853]" ' vla:2616
            Exit Function
        Case 469
            vla_step_text = "Format cell F43 as a number with 0 decimals and thousands separators. [line 854]" ' vla:2617
            Exit Function
        Case 470
            vla_step_text = "Put 1234.56 into cell F44. [line 855]" ' vla:2618
            Exit Function
        Case 471
            vla_step_text = "Format cell F44 as dollars. [line 856]" ' vla:2619
            Exit Function
        Case 472
            vla_step_text = "Put 1234.56 into cell F45. [line 857]" ' vla:2620
            Exit Function
        Case 473
            vla_step_text = "Format cell F45 as euros with 0 decimals. [line 858]" ' vla:2621
            Exit Function
        Case 474
            vla_step_text = "Put 1234.56 into cell F46. [line 859]" ' vla:2622
            Exit Function
        Case 475
            vla_step_text = "Format cell F46 as pounds. [line 860]" ' vla:2623
            Exit Function
        Case 476
            vla_step_text = "Put 1234.56 into cell F47. [line 861]" ' vla:2624
            Exit Function
        Case 477
            vla_step_text = "Format cell F47 as accounting in dollars. [line 862]" ' vla:2625
            Exit Function
        Case 478
            vla_step_text = "Put -1234.56 into cell F48. [line 863]" ' vla:2626
            Exit Function
        Case 479
            vla_step_text = "Format cell F48 as accounting in euros with 0 decimals. [line 864]" ' vla:2627
            Exit Function
        Case 480
            vla_step_text = "Put 0.125 into cell F49. [line 865]" ' vla:2628
            Exit Function
        Case 481
            vla_step_text = "Format range F49 as percent. [line 866]" ' vla:2629
            Exit Function
        Case 482
            vla_step_text = "Put 0.125 into cell F50. [line 867]" ' vla:2630
            Exit Function
        Case 483
            vla_step_text = "Format cell F50 as percent with 2 decimals. [line 868]" ' vla:2631
            Exit Function
        Case 484
            vla_step_text = "Put 46000 into cell F51. [line 869]" ' vla:2632
            Exit Function
        Case 485
            vla_step_text = "Format cell F51 as a short date. [line 870]" ' vla:2633
            Exit Function
        Case 486
            vla_step_text = "Put 46000 into cell F52. [line 871]" ' vla:2634
            Exit Function
        Case 487
            vla_step_text = "Format cell F52 as a long date. [line 872]" ' vla:2635
            Exit Function
        Case 488
            vla_step_text = "Put 46000 into cell F53. [line 873]" ' vla:2636
            Exit Function
        Case 489
            vla_step_text = "Format cell F53 as an ISO date. [line 874]" ' vla:2637
            Exit Function
        Case 490
            vla_step_text = "Put 0.5625 into cell F54. [line 875]" ' vla:2638
            Exit Function
        Case 491
            vla_step_text = "Format cell F54 as a time. [line 876]" ' vla:2639
            Exit Function
        Case 492
            vla_step_text = "Put 42 into cell F55. [line 877]" ' vla:2640
            Exit Function
        Case 493
            vla_step_text = "Format cell F55 as text for new entries. [line 878]" ' vla:2641
            Exit Function
        Case 494
            vla_step_text = "Put 1234.56 into cell F56. [line 879]" ' vla:2642
            Exit Function
        Case 495
            vla_step_text = "Format cell F56 as dollars. [line 880]" ' vla:2643
            Exit Function
        Case 496
            vla_step_text = "Format cell F56 as general. [line 881]" ' vla:2644
            Exit Function
        Case 497
            vla_step_text = "Put 42 into cell F57. [line 882]" ' vla:2645
            Exit Function
        Case 498
            vla_step_text = "Format cell F57 using ""00000"". [line 883]" ' vla:2646
            Exit Function
        Case 499
            vla_step_text = "Add a border colored ""#FF0000"" around range H40:J42. [line 889]" ' vla:2647
            Exit Function
        Case 500
            vla_step_text = "Add a bottom border colored green to range H44:J44. [line 890]" ' vla:2648
            Exit Function
        Case 501
            vla_step_text = "Add borders colored blue to every cell in range H46:J48. [line 891]" ' vla:2649
            Exit Function
        Case 502
            vla_step_text = "Go to sheet Output. [line 893]" ' vla:2650
            Exit Function
        Case 503
            vla_step_text = "Tidy-up. [line 895]" ' vla:2651
            Exit Function
        Case 504
            vla_step_text = "Turn on screen updating. [line 896]" ' vla:2652
            Exit Function
        Case 505
            vla_step_text = "Log ""report finished"". [line 897]" ' vla:2653
            Exit Function
        Case Else
            vla_step_text = "an unknown step" ' vla:2654
            Exit Function
    End Select
End Function


