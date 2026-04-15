// ================================================================
//  SCHOOL MANAGEMENT SYSTEM — Code.gs v5.1
//  FIXES: Security, Performance, Data Integrity, Scalability
// ================================================================
var SH = {
  USERS:'Users', STUDENTS:'Students', TEACHERS:'Teachers',
  CLASSES:'Classes', SECTIONS:'Sections', SUBJECTS:'Subjects',
  EXAMS:'Exams', MARKS:'Marks', ATTENDANCE:'Attendance',
  RESULTS:'Results', SESSIONS:'AcademicSessions',
  TIMETABLE:'Timetable', GRADING:'GradingSystem',
  RESULT_PUBLISH:'ResultPublish',
  NOTIFICATIONS:'Notifications',
  AUDIT:'AuditLog', SETTINGS:'SchoolSettings',
  // Fee Management v2
  FEE_STRUCTURE:'FeeStructure',
  PAYMENTS:'Payments'
};

// ── Module-level SpreadsheetApp cache (saves repeated service lookups) ──
var _ss = null;
function getSS() {
  if (!_ss) _ss = SpreadsheetApp.getActiveSpreadsheet();
  return _ss;
}

// ── Per-request in-memory sheet cache (avoid redundant full scans) ──
var _sheetCache = {};
function getCachedSheet(ss, cache, shName, ttlSeconds) {
  // Support both old single-arg and new 4-arg call styles
  if (typeof ss === 'string') {
    // Old-style: getCachedSheet(shName)
    shName = ss; ss = getSS(); cache = CacheService.getScriptCache(); ttlSeconds = 180;
  }
  var key = 'sheet_' + String(shName).replace(/[^a-zA-Z0-9]/g,'_');
  var hit = null;
  try { hit = cache.get(key); } catch(e){}
  if (hit) { try { return JSON.parse(hit); } catch(e){} }
  var data = parseSheet(ss, shName);
  try { cache.put(key, JSON.stringify(data), ttlSeconds||180); } catch(e){}
  return data;
}
function clearSheetCache(shName) {
  try{
    var cache = CacheService.getScriptCache();
    if(shName){
      cache.remove('sheet_'+shName.replace(/[^a-zA-Z0-9]/g,'_'));
    } else {
      cache.removeAll(['alldata_settings','inst_schedule_v1']);
    }
  }catch(e){}
}
function doGet(e) {
  var output = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('EduManage Pro')
    .addMetaTag('viewport','width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no');
  // setXFrameOptionsMode: use try/catch — enum resolution fails in some GAS runtimes
  try {
    output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.SAMEORIGIN);
  } catch(e) {
    try { output.setXFrameOptionsMode(1); } catch(e2) {} // 1 = SAMEORIGIN numeric value
  }
  return output;
}

// ═══════════════════════════════════════════════════════════
//  PASSWORD HELPERS (SHA-256 — no plain text storage)
// ═══════════════════════════════════════════════════════════
function hashPassword(plain) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(plain),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b){
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('');
}
function checkPassword(plain, hashed) {
  // Support both legacy plain-text and new hashed passwords during migration
  if (hashed.length === 64) return hashPassword(plain) === hashed; // SHA-256 hex
  return plain === hashed; // legacy plain — triggers upgrade on next save
}

// ═══════════════════════════════════════════════════════════
//  SETUP
// ═══════════════════════════════════════════════════════════
function setupSystem() {
  var ss = getSS();
  // Safety guard: warn if any core sheet already has live data
  var liveSheets = [SH.STUDENTS, SH.FEES, SH.MARKS, SH.ATTENDANCE].filter(function(name) {
    var sh = ss.getSheetByName(name);
    return sh && sh.getLastRow() > 1;
  });
  if (liveSheets.length > 0) {
    var ui = SpreadsheetApp.getUi();
    var resp = ui.alert(
      '⚠️ LIVE DATA DETECTED',
      'These sheets have existing data:\n' + liveSheets.join(', ') + '\n\nRunning setup will ERASE ALL DATA.\nUse "Run Setup (New Sheets Only)" to safely add v5.0 sheets.\n\nType YES to confirm full reset, or Cancel to abort.',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp !== ui.Button.OK) return;
  }
  mkSheet(ss,SH.USERS,['UserID','Email','Password','Role','Name','Avatar','AssociatedID','Status','LastLogin','CreatedAt','MustChangePassword','ResetOTP','OTPExpiry'],'#1a73e8');
  mkSheet(ss,SH.STUDENTS,['StudentID','RollNo','AdmissionYear','FirstName','LastName','Email','Gender','DOB','BloodGroup','ClassID','SectionID','Phone','Address','GuardianName','GuardianPhone','GuardianEmail','AdmissionDate','Status','PhotoURL'],'#34A853');
  mkSheet(ss,SH.TEACHERS,['TeacherID','EmpID','FirstName','LastName','Email','Gender','Phone','Qualification','Specialization','JoinDate','AssignedClasses','AssignedSubjects','Status'],'#FBBC04');
  mkSheet(ss,SH.CLASSES,['ClassID','ClassName','ClassTeacherID','AcademicYear','Capacity','Status'],'#EA4335');
  mkSheet(ss,SH.SECTIONS,['SectionID','ClassID','SectionName','Room','Capacity'],'#9C27B0');
  mkSheet(ss,SH.SUBJECTS,['SubjectID','SubjectName','SubjectCode','ClassID','TeacherID','MaxMarks','PassMarks','Type'],'#00BCD4');
  mkSheet(ss,SH.EXAMS,['ExamID','ExamName','ExamType','ClassID','AcademicYear','StartDate','EndDate','TotalMarks','Published','CreatedAt'],'#795548');
  mkSheet(ss,SH.MARKS,['MarkID','ExamID','StudentID','SubjectID','ClassID','TeacherID','MarksObtained','MaxMarks','Grade','Remarks','CreatedAt','UpdatedAt'],'#607D8B');
  mkSheet(ss,SH.ATTENDANCE,['AttID','Date','StudentID','ClassID','SectionID','Status','Remarks','MarkedBy','MarkedAt'],'#009688');
  mkSheet(ss,SH.RESULTS,['ResultID','ExamID','StudentID','TotalMarks','MarksObtained','Percentage','Grade','Rank','Published','GeneratedAt'],'#FF5722');
  mkSheet(ss,SH.SESSIONS,['SessionID','SessionName','StartDate','EndDate','IsActive'],'#3F51B5');
  mkSheet(ss,SH.TIMETABLE,['TimetableID','ClassID','SectionID','Day','Period','TimeFrom','TimeTo','SubjectID','TeacherID','Room'],'#E91E63');
  mkSheet(ss,SH.GRADING,['GradeID','GradeName','MinPercent','MaxPercent','GradePoint','Remarks'],'#FF9800');
  mkSheet(ss,SH.RESULT_PUBLISH,['PublishID','ExamID','ClassID','ExamType','PublishStatus','PublishedBy','PublishedAt'],'#2196F3');
  mkSheet(ss,SH.FEE_STRUCTURE,['FeeID','FeeName','Category','Amount','ClassID','Frequency','AppliesTo','AcademicYear','Description','CreatedAt'],'#4CAF50');
  mkSheet(ss,SH.FEES,['PaymentID','StudentID','ClassID','FeeStructureID','FeeName','Month','FeeAmount','PaidAmount','DueAmount','PaymentDate','CollectedBy','Status','Receipt','Notes','CreatedAt'],'#8BC34A');
  mkSheet(ss,SH.NOTIFICATIONS,['NotifID','Title','Message','SenderID','SenderRole','SenderName','TargetType','TargetID','IsRead','CreatedAt'],'#FF5722');
  mkSheet(ss,SH.AUDIT,['AuditID','ActionType','RecordID','TableName','PerformedBy','Role','Timestamp','OldValue','NewValue','Description'],'#607D8B');
  mkSheet(ss,SH.SETTINGS,['Key','Value','UpdatedBy','UpdatedAt'],'#9E9E9E');
  // v5.0 new sheets
  mkSheet(ss,SH.FEE_HEADS,['FeeHeadID','FeeName','Frequency','Description','IsActive','CreatedAt'],'#4DB6AC');
  mkSheet(ss,SH.PAYMENTS,['PaymentID','StudentID','PaymentDate','PaymentMode','ReceiptNumber','TotalAmount','CollectedBy','Notes','CreatedAt'],'#26A69A');
  mkSheet(ss,SH.PAYMENT_ITEMS,['ItemID','PaymentID','FeeHeadID','FeeHeadName','Amount','Month','AcademicYear'],'#80CBC4');
  mkSheet(ss,SH.DISCOUNTS,['DiscountID','StudentID','FeeHeadID','FeeHeadName','DiscountPercent','Reason','AcademicYear','CreatedBy','CreatedAt'],'#A5D6A7');
  mkSheet(ss,SH.LATE_FEE_RULES,['RuleID','DaysLate','LateFeeAmount','Description','IsActive','CreatedAt'],'#FFAB91');
  seedData(ss);
  SpreadsheetApp.getUi().alert('EduManage Pro v5.1 Setup Complete!\n\nadmin@school.edu / admin123\nteacher@school.edu / teacher123\nstudent@school.edu / student123\n\nPasswords are stored hashed. Please change defaults after first login.');
}

function mkSheet(ss,name,headers,color){
  var sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name); else sh.clearContents();
  sh.setTabColor(color);
  var r=sh.getRange(1,1,1,headers.length);
  r.setValues([headers]);
  r.setBackground(color).setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);
}
function aR(ss,s,v){ss.getSheetByName(s).appendRow(v);}

function seedData(ss){
  var n=new Date().toISOString();
  // Sessions
  aR(ss,SH.SESSIONS,['SES001','2025-2026','2025-06-01','2026-04-30','TRUE']);
  // Grading
  [['GRD001','A+',97,100,4.0,'Outstanding'],['GRD002','A',93,96,4.0,'Excellent'],
   ['GRD003','A-',90,92,3.7,'Excellent'],['GRD004','B+',87,89,3.3,'Very Good'],
   ['GRD005','B',83,86,3.0,'Good'],['GRD006','B-',80,82,2.7,'Good'],
   ['GRD007','C+',77,79,2.3,'Above Average'],['GRD008','C',73,76,2.0,'Average'],
   ['GRD009','C-',70,72,1.7,'Average'],['GRD010','D',60,69,1.0,'Below Average'],
   ['GRD011','F',0,59,0,'Fail']
  ].forEach(function(g){aR(ss,SH.GRADING,g);});
  // Users — passwords stored as SHA-256 hash
  // Users: UserID,Email,Password,Role,Name,Avatar,AssociatedID,Status,LastLogin,CreatedAt,MustChangePassword,ResetOTP,OTPExpiry
  aR(ss,SH.USERS,['USR001','admin@school.edu',hashPassword('admin123'),'admin','Dr. Sarah Mitchell','SM','ADM001','Active','',n,'FALSE','','']);
  aR(ss,SH.USERS,['USR002','teacher@school.edu',hashPassword('teacher123'),'teacher','Mr. James Wilson','JW','TCH001','Active','',n,'FALSE','','']);
  aR(ss,SH.USERS,['USR003','teacher2@school.edu',hashPassword('teacher123'),'teacher','Ms. Priya Sharma','PS','TCH002','Active','',n,'FALSE','','']);
  aR(ss,SH.USERS,['USR004','student@school.edu',hashPassword('student123'),'student','Emily Chen','EC','STU001','Active','',n,'FALSE','','']);
  aR(ss,SH.USERS,['USR005','student2@school.edu',hashPassword('student123'),'student','Marcus Thompson','MT','STU002','Active','',n,'FALSE','','']);
  // Classes
  aR(ss,SH.CLASSES,['CLS001','Grade 10','TCH001','2025-2026',35,'Active']);
  aR(ss,SH.CLASSES,['CLS002','Grade 11','TCH002','2025-2026',35,'Active']);
  aR(ss,SH.CLASSES,['CLS003','Grade 12','TCH001','2025-2026',35,'Active']);
  // Sections
  aR(ss,SH.SECTIONS,['SEC001','CLS001','Section A','Room 101',35]);
  aR(ss,SH.SECTIONS,['SEC002','CLS001','Section B','Room 102',35]);
  // Teachers
  aR(ss,SH.TEACHERS,['TCH001','EMP-001','James','Wilson','teacher@school.edu','Male','09301234567','M.Sc Mathematics','Mathematics','2023-06-01','CLS001,CLS002','SUBJ001','Active']);
  aR(ss,SH.TEACHERS,['TCH002','EMP-002','Priya','Sharma','teacher2@school.edu','Female','09311234567','M.Sc Physics','Science','2023-06-01','CLS001','SUBJ002','Active']);
  aR(ss,SH.TEACHERS,['TCH003','EMP-003','Robert','Adams','robert@school.edu','Male','09321234567','M.A. English','English','2022-06-01','CLS001','SUBJ003','Active']);
  aR(ss,SH.TEACHERS,['TCH004','EMP-004','Ayesha','Khan','ayesha@school.edu','Female','09331234567','M.A. History','Social Studies','2021-06-01','CLS001','SUBJ004','Active']);
  aR(ss,SH.TEACHERS,['TCH005','EMP-005','Daniel','Choi','daniel@school.edu','Male','09341234567','B.P.E.','Physical Ed','2024-06-01','CLS001','SUBJ005','Active']);
  // Subjects
  aR(ss,SH.SUBJECTS,['SUBJ001','Mathematics','MATH','CLS001','TCH001',100,40,'Theory']);
  aR(ss,SH.SUBJECTS,['SUBJ002','Science','SCI','CLS001','TCH002',100,40,'Theory']);
  aR(ss,SH.SUBJECTS,['SUBJ003','English','ENG','CLS001','TCH003',100,40,'Theory']);
  aR(ss,SH.SUBJECTS,['SUBJ004','Social Studies','SOC','CLS001','TCH004',100,40,'Theory']);
  aR(ss,SH.SUBJECTS,['SUBJ005','Physical Ed','PE','CLS001','TCH005',50,20,'Practical']);
  // Students — now includes AdmissionYear column (position 2, after RollNo)
  aR(ss,SH.STUDENTS,['STU001','2025-0001',2025,'Emily','Chen','student@school.edu','Female','2009-04-15','B+','CLS001','SEC001','09171234567','123 Maple St','Lin Chen','09171234568','lchen@email.com','2025-06-01','Active','']);
  aR(ss,SH.STUDENTS,['STU002','2025-0002',2025,'Marcus','Thompson','student2@school.edu','Male','2009-07-22','O+','CLS001','SEC001','09181234567','456 Oak Ave','Dale Thompson','09181234568','dale@email.com','2025-06-01','Active','']);
  aR(ss,SH.STUDENTS,['STU003','2025-0003',2025,'Sophia','Rodriguez','sophia@school.edu','Female','2009-11-08','A+','CLS001','SEC001','09191234567','789 Pine Rd','Ana Rodriguez','09191234568','ana@email.com','2025-06-01','Active','']);
  aR(ss,SH.STUDENTS,['STU004','2025-0004',2025,'Aiden','Park','aiden@school.edu','Male','2010-02-14','AB+','CLS001','SEC002','09201234567','321 Elm St','Jin Park','09201234568','jin@email.com','2025-06-01','Active','']);
  aR(ss,SH.STUDENTS,['STU005','2025-0005',2025,'Isabella','Torres','isabella@school.edu','Female','2009-09-30','O-','CLS001','SEC001','09211234567','654 Cedar Ln','Maria Torres','09211234568','mtorres@email.com','2025-06-01','Active','']);
  // Exams
  aR(ss,SH.EXAMS,['EXM001','Unit Test 1','unit','CLS001','2025-2026','2025-08-01','2025-08-05',450,'TRUE',n]);
  aR(ss,SH.EXAMS,['EXM002','Mid Term','midterm','CLS001','2025-2026','2025-09-15','2025-09-20',450,'TRUE',n]);
  aR(ss,SH.EXAMS,['EXM003','Half Yearly','halfyearly','CLS001','2025-2026','2025-11-01','2025-11-05',450,'FALSE',n]);
  // Marks
  var marks=[
    ['MRK001','EXM001','STU001','SUBJ001','CLS001','TCH001',88,100,'B+','',n,n],
    ['MRK002','EXM001','STU001','SUBJ002','CLS001','TCH002',91,100,'A-','',n,n],
    ['MRK003','EXM001','STU001','SUBJ003','CLS001','TCH003',85,100,'B+','',n,n],
    ['MRK004','EXM001','STU001','SUBJ004','CLS001','TCH004',79,100,'C+','',n,n],
    ['MRK005','EXM001','STU001','SUBJ005','CLS001','TCH005',44,50,'A','',n,n],
    ['MRK006','EXM001','STU002','SUBJ001','CLS001','TCH001',72,100,'C','',n,n],
    ['MRK007','EXM001','STU002','SUBJ002','CLS001','TCH002',68,100,'C-','',n,n],
    ['MRK008','EXM001','STU002','SUBJ003','CLS001','TCH003',75,100,'C+','',n,n],
    ['MRK009','EXM001','STU002','SUBJ004','CLS001','TCH004',80,100,'B-','',n,n],
    ['MRK010','EXM001','STU002','SUBJ005','CLS001','TCH005',38,50,'B-','',n,n],
    ['MRK011','EXM001','STU003','SUBJ001','CLS001','TCH001',95,100,'A','',n,n],
    ['MRK012','EXM001','STU003','SUBJ002','CLS001','TCH002',93,100,'A','',n,n],
    ['MRK013','EXM001','STU003','SUBJ003','CLS001','TCH003',97,100,'A+','',n,n],
    ['MRK014','EXM001','STU003','SUBJ004','CLS001','TCH004',91,100,'A-','',n,n],
    ['MRK015','EXM001','STU003','SUBJ005','CLS001','TCH005',48,50,'A+','',n,n]
  ];
  marks.forEach(function(m){aR(ss,SH.MARKS,m);});
  // Results
  aR(ss,SH.RESULTS,['RES001','EXM001','STU001',450,387,86.0,'B+',2,'TRUE',n]);
  aR(ss,SH.RESULTS,['RES002','EXM001','STU002',450,333,74.0,'C+',3,'TRUE',n]);
  aR(ss,SH.RESULTS,['RES003','EXM001','STU003',450,424,94.2,'A',1,'TRUE',n]);
  // Publish
  aR(ss,SH.RESULT_PUBLISH,['PUB001','EXM001','CLS001','unit','TRUE','ADM001',n]);
  // Fee Structure
  aR(ss,SH.FEE_STRUCTURE,['FEE_STR001','Tuition Fee','CLS001',5000,'2025-08-10','2025-2026','Monthly tuition fee',n]);
  aR(ss,SH.FEE_STRUCTURE,['FEE_STR002','Development Fee','CLS001',2000,'2025-08-10','2025-2026','Annual development fee',n]);
  aR(ss,SH.FEE_STRUCTURE,['FEE_STR003','Library Fee','CLS001',500,'2025-08-10','2025-2026','Annual library fee',n]);
  // Fees
  aR(ss,SH.FEES,['PAY001','STU001','CLS001','FEE_STR001','Tuition Fee','August',5000,5000,0,'2025-08-05','ADM001','Paid','RCP001','',n]);
  aR(ss,SH.FEES,['PAY002','STU001','CLS001','FEE_STR001','Tuition Fee','September',5000,3000,2000,'2025-09-10','ADM001','Partial','RCP002','Partial payment',n]);
  aR(ss,SH.FEES,['PAY003','STU002','CLS001','FEE_STR001','Tuition Fee','August',5000,0,5000,'','','Pending','','',n]);
  // Notifications
  aR(ss,SH.NOTIFICATIONS,['NTF001','Welcome to EduManage Pro','System has been successfully set up. Welcome!','ADM001','admin','Dr. Sarah Mitchell','all','all','FALSE',n]);
  aR(ss,SH.NOTIFICATIONS,['NTF002','Unit Test Results Published','Unit Test 1 results are now available. Check your dashboard.','ADM001','admin','Dr. Sarah Mitchell','all','all','FALSE',n]);
  aR(ss,SH.NOTIFICATIONS,['NTF003','Fee Reminder','August tuition fee is due by 10th. Please pay on time.','ADM001','admin','Dr. Sarah Mitchell','class','CLS001','FALSE',n]);
  // Settings
  [['school_name','Excel Academy','ADM001',n],
   ['school_address','123 Education Lane, Knowledge City','ADM001',n],
   ['school_phone','+1-555-0100','ADM001',n],
   ['school_email','info@excelacademy.edu','ADM001',n],
   ['school_logo','','ADM001',n],
   ['show_photo_marksheet','true','ADM001',n],
   ['academic_year','2025-2026','ADM001',n],
   ['principal_name','Dr. Sarah Mitchell','ADM001',n],
   ['school_motto','Excellence in Education','ADM001',n],
   ['school_color','#6366f1','ADM001',n]
  ].forEach(function(s){aR(ss,SH.SETTINGS,s);});
  // Attendance
  ['2025-08-04','2025-08-05','2025-08-06','2025-08-07','2025-08-08'].forEach(function(d,i){
    aR(ss,SH.ATTENDANCE,['ATT0'+i,d,'STU001','CLS001','SEC001',i===2?'A':'P','','TCH001',n]);
  });
}

// ═══════════════════════════════════════════════════════════
//  CORE HELPERS
// ═══════════════════════════════════════════════════════════
function sheetToObjects(shName){
  return parseSheet(getSS(), shName);
}

function upsertRow(shName,idField,data){
  var ss=getSS();
  var sh=ss.getSheetByName(shName);
  if(!sh) return {success:false,message:'Sheet not found: '+shName};
  var lc=sh.getLastColumn(); if(lc<1) return {success:false};
  var lr=sh.getLastRow();
  var allVals=sh.getRange(1,1,Math.max(lr,1),lc).getValues();
  var headers=allVals[0];

  // Auto-add any missing columns that exist in data but not in sheet header
  Object.keys(data).forEach(function(key){
    if(key && headers.indexOf(key) < 0){
      var newCol = headers.length + 1;
      sh.getRange(1, newCol).setValue(key);
      headers.push(key);
      lc = headers.length;
    }
  });

  var idIdx=headers.indexOf(idField);
  if(data[idField]&&lr>=2){
    for(var i=1;i<allVals.length;i++){
      if(String(allVals[i][idIdx])===String(data[idField])){
        var finalRow=headers.map(function(h,k){
          return data[h]!==undefined ? data[h] : (allVals[i][k]!==undefined ? allVals[i][k] : '');
        });
        sh.getRange(i+1,1,1,lc).setValues([finalRow]);
        SpreadsheetApp.flush();
        clearSheetCache(shName);
        return {success:true,action:'updated',id:data[idField]};
      }
    }
  }
  if(!data[idField]) data[idField]=shName.substr(0,3).toUpperCase()+'_'+Date.now();
  var newRow=headers.map(function(h){return data[h]!==undefined?data[h]:'';});
  sh.appendRow(newRow);
  SpreadsheetApp.flush();
  clearSheetCache(shName);
  return {success:true,action:'created',id:data[idField]};
}

function removeRow(shName,idField,idVal){
  var ss=getSS();
  var sh=ss.getSheetByName(shName);
  if(!sh) return {success:false};
  var lr=sh.getLastRow(); if(lr<2) return {success:false};
  var vals=sh.getRange(2,1,lr-1,sh.getLastColumn()).getValues();
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var idIdx=headers.indexOf(idField);
  for(var i=0;i<vals.length;i++){
    if(String(vals[i][idIdx])===String(idVal)){
      sh.deleteRow(i+2);
      SpreadsheetApp.flush();
      clearSheetCache(shName);
      return {success:true};
    }
  }
  return {success:false};
}

// Audit log helper — no flush (batched at end of main operation)
var _auditBuffer = [];
function addAudit(actionType,recordId,tableName,performedBy,role,oldVal,newVal,desc){
  try {
    var ss=getSS();
    var sh=ss.getSheetByName(SH.AUDIT);
    if(!sh) return;
    // Auto-create headers if sheet is empty
    if(sh.getLastRow()===0){
      sh.appendRow(['AuditID','ActionType','RecordID','TableName','PerformedBy','Role',
                    'Timestamp','OldValue','NewValue','Description']);
    }
    sh.appendRow([
      'AUD_'+Date.now()+'_'+Math.random().toString(36).slice(2,6),
      String(actionType||''),
      String(recordId||''),
      String(tableName||''),
      String(performedBy||'system'),
      String(role||'system'),
      new Date().toISOString(),
      oldVal ? JSON.stringify(oldVal).substring(0,500) : '',
      newVal ? JSON.stringify(newVal).substring(0,500) : '',
      String(desc||'')
    ]);
  }catch(e){}
}

// ═══════════════════════════════════════════════════════════
//  AUTH
// ═══════════════════════════════════════════════════════════
function login(email,password){
  try {
    var ss=getSS(); var sh=ss.getSheetByName(SH.USERS);
    var data=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getValues();
    var headers=data[0];
    var ei=headers.indexOf('Email'),pi=headers.indexOf('Password'),
        ri=headers.indexOf('Role'),ni=headers.indexOf('Name'),
        ai=headers.indexOf('Avatar'),si=headers.indexOf('Status'),
        ii=headers.indexOf('AssociatedID'),ui=headers.indexOf('UserID'),
        li=headers.indexOf('LastLogin');
    for(var i=1;i<data.length;i++){
      var row=data[i];
      if(String(row[ei]).trim()!==String(email).trim()) continue;
      if(String(row[si]).trim()!=='Active') continue;
      if(!checkPassword(String(password).trim(), String(row[pi]).trim())) continue;
      // Upgrade legacy plain-text password to hash on successful login
      if(String(row[pi]).length!==64){
        sh.getRange(i+1,pi+1).setValue(hashPassword(password));
      }
      sh.getRange(i+1,li+1).setValue(new Date().toISOString());
      SpreadsheetApp.flush();
      return {success:true,user:{
        id:String(row[ui]),email:String(row[ei]),
        role:String(row[ri]).trim().toLowerCase(),
        name:String(row[ni]),avatar:String(row[ai]),
        assocId:String(row[ii])
      }};
    }
    return {success:false,message:'Invalid email or password.'};
  }catch(e){return {success:false,message:'Error: '+e.message};}
}

// changePassword: requires current password — proper authenticated flow
function changePassword(email, currentPass, newPass) {
  try {
    var ss=getSS(); var sh=ss.getSheetByName(SH.USERS);
    var data=sh.getDataRange().getValues(); var headers=data[0];
    var ei=headers.indexOf('Email'),pi=headers.indexOf('Password'),si=headers.indexOf('Status');
    for(var i=1;i<data.length;i++){
      if(String(data[i][ei]).trim()!==String(email).trim()) continue;
      if(String(data[i][si]).trim()!=='Active') return {success:false,message:'Account inactive'};
      if(!checkPassword(String(currentPass).trim(), String(data[i][pi]).trim()))
        return {success:false,message:'Current password is incorrect.'};
      if(!newPass||newPass.length<6) return {success:false,message:'New password must be at least 6 characters.'};
      sh.getRange(i+1,pi+1).setValue(hashPassword(newPass));
      SpreadsheetApp.flush();
      return {success:true};
    }
    return {success:false,message:'Email not found.'};
  }catch(e){return {success:false,message:e.message};}
}

// adminResetPassword: admin sets a new temporary password (does NOT require current password)
// Security: should only be callable after verifying the caller is admin via session token
function adminResetPassword(email, newPass, adminId) {
  try {
    if (!adminId) return {success:false, message:'Unauthorized'};
    if (!newPass || newPass.length < 6) return {success:false, message:'Password too short (min 6 chars)'};
    var ss=getSS(); var sh=ss.getSheetByName(SH.USERS);
    var data=sh.getDataRange().getValues(); var headers=data[0];
    var ei=headers.indexOf('Email'),pi=headers.indexOf('Password');
    for(var i=1;i<data.length;i++){
      if(String(data[i][ei]).trim()===String(email).trim()){
        sh.getRange(i+1,pi+1).setValue(hashPassword(newPass));
        SpreadsheetApp.flush();
        addAudit('PASSWORD_RESET',email,SH.USERS,adminId,'admin',null,null,'Admin reset password for '+email);
        SpreadsheetApp.flush();
        return {success:true};
      }
    }
    return {success:false,message:'Email not found.'};
  }catch(e){return {success:false,message:e.message};}
}

// ═══════════════════════════════════════════════════════════
//  ALL DATA BULK LOAD
// ═══════════════════════════════════════════════════════════
// Fast helper: parse one sheet's data given already-fetched values
// Known field aliases → canonical name
var _FIELD_MAP = {
  // Students
  'studentid':'StudentID','student_id':'StudentID',
  'firstname':'FirstName','first_name':'FirstName','fname':'FirstName',
  'lastname':'LastName','last_name':'LastName','lname':'LastName','surname':'LastName',
  'name':'Name',
  'email':'Email','emailaddress':'Email','email_address':'Email',
  'rollno':'RollNo','roll_no':'RollNo','rollnumber':'RollNo','admissionno':'RollNo',
  'admissionyear':'AdmissionYear','admission_year':'AdmissionYear',
  'classid':'ClassID','class_id':'ClassID',
  'sectionid':'SectionID','section_id':'SectionID',
  'gender':'Gender','sex':'Gender',
  'dob':'DOB','dateofbirth':'DOB','date_of_birth':'DOB','birthdate':'DOB',
  'bloodgroup':'BloodGroup','blood_group':'BloodGroup','blood':'BloodGroup',
  'phone':'Phone','mobile':'Phone','contact':'Phone','phonenumber':'Phone',
  'address':'Address',
  'guardianname':'GuardianName','guardian_name':'GuardianName','parentname':'GuardianName',
  'guardianphone':'GuardianPhone','guardian_phone':'GuardianPhone','parentphone':'GuardianPhone',
  'guardianemail':'GuardianEmail','guardian_email':'GuardianEmail','parentemail':'GuardianEmail',
  'status':'Status',
  'photourl':'PhotoURL','photo':'PhotoURL','photo_url':'PhotoURL',
  // Teachers
  'teacherid':'TeacherID','teacher_id':'TeacherID',
  'empid':'EmpID','emp_id':'EmpID','employeeid':'EmpID',
  'specialization':'Specialization','specialisation':'Specialization','subject':'Specialization',
  'qualification':'Qualification',
  'joindate':'JoinDate','join_date':'JoinDate','joiningdate':'JoinDate',
  'assignedclasses':'AssignedClasses','assigned_classes':'AssignedClasses',
  'assignedsubjects':'AssignedSubjects','assigned_subjects':'AssignedSubjects',
  // Classes
  'classid':'ClassID','classname':'ClassName','class_name':'ClassName',
  'classteacherid':'ClassTeacherID','class_teacher_id':'ClassTeacherID',
  'academicyear':'AcademicYear','academic_year':'AcademicYear',
  // Common
  'createdat':'CreatedAt','created_at':'CreatedAt',
  'updatedat':'UpdatedAt','updated_at':'UpdatedAt',
  // Settings sheet
  'key':'Key','value':'Value','updatedby':'UpdatedBy',
  // Exams
  'examid':'ExamID','exam_id':'ExamID',
  'examname':'ExamName','exam_name':'ExamName','examnm':'ExamName',
  'examtype':'ExamType','exam_type':'ExamType',
  'classid':'ClassID','class_id':'ClassID',
  // Marks
  'markid':'MarkID','mark_id':'MarkID','marksid':'MarkID',
  'marksobtained':'MarksObtained','marks_obtained':'MarksObtained','marks':'MarksObtained','score':'MarksObtained',
  'maxmarks':'MaxMarks','max_marks':'MaxMarks','totalmarks':'MaxMarks',
  'grade':'Grade',
  // Attendance
  'attid':'AttID','att_id':'AttID',
  'date':'Date',
  'markedby':'MarkedBy','marked_by':'MarkedBy',
  // Fees
  'paymentid':'PaymentID','payment_id':'PaymentID',
  'feeamount':'FeeAmount','fee_amount':'FeeAmount','amount':'FeeAmount',
  'paidamount':'PaidAmount','paid_amount':'PaidAmount','paid':'PaidAmount',
  'dueamount':'DueAmount','due_amount':'DueAmount','due':'DueAmount',
  'paymentdate':'PaymentDate','payment_date':'PaymentDate',
  'paymentmode':'PaymentMode','payment_mode':'PaymentMode','mode':'PaymentMode',
  'collectedby':'CollectedBy','collected_by':'CollectedBy',
  // Results
  'resultid':'ResultID','result_id':'ResultID',
  'percentage':'Percentage','percent':'Percentage',
  'rank':'Rank','position':'Rank',
  'published':'Published',
  // Notifications
  'notifid':'NotifID','notif_id':'NotifID',
  'title':'Title','message':'Message','isread':'IsRead','is_read':'IsRead',
  'senderid':'SenderID','sender_id':'SenderID',
  'targettype':'TargetType','target_type':'TargetType',
  'targetid':'TargetID','target_id':'TargetID',
  // Audit log
  'auditid':'AuditID','audit_id':'AuditID',
  'actiontype':'ActionType','action_type':'ActionType','action':'ActionType',
  'recordid':'RecordID','record_id':'RecordID',
  'tablename':'TableName','table_name':'TableName','table':'TableName','sheetname':'TableName',
  'performedby':'PerformedBy','performed_by':'PerformedBy','userid':'PerformedBy',
  'timestamp':'Timestamp','time':'Timestamp','datetime':'Timestamp',
  'oldvalue':'OldValue','old_value':'OldValue','oldval':'OldValue','old_val':'OldValue',
  'newvalue':'NewValue','new_value':'NewValue','newval':'NewValue','new_val':'NewValue',
  'description':'Description','desc':'Description','details':'Description','notes':'Description',
};

function normalizeKey(h) {
  var s = String(h || '').trim();
  if (!s) return '';
  // Check lookup table first (case-insensitive, no spaces/underscores)
  var lookup = s.toLowerCase().replace(/[\s_-]+/g,'');
  if (_FIELD_MAP[lookup]) return _FIELD_MAP[lookup];
  // Fallback: convert "First Name" / "first_name" → camelCase
  var converted = s.replace(/[\s_]+(.)/g, function(_, c){ return c.toUpperCase(); });
  return converted;
}

// Post-process a student record to fill in missing fields from alternatives
function normalizeStudentRecord(s) {
  if (!s) return s;
  // Normalize all string values - trim whitespace, convert 'undefined' string to empty
  Object.keys(s).forEach(function(k){
    if (typeof s[k] === 'string') {
      s[k] = s[k].trim();
      if (s[k] === 'undefined' || s[k] === 'null') s[k] = '';
    }
  });
  // If FirstName is missing, try to split a "Name" field
  if (!s.FirstName && !s.LastName && s.Name) {
    var parts = String(s.Name).trim().split(/\s+/);
    s.FirstName = parts[0] || '';
    s.LastName  = parts.slice(1).join(' ') || '';
  }
  // If still no name, use email prefix
  if (!s.FirstName && s.Email) {
    var ep = String(s.Email).split('@')[0].replace(/[^a-zA-Z]/g,' ').trim();
    var parts2 = ep.split(/\s+/);
    s.FirstName = parts2[0] ? (parts2[0].charAt(0).toUpperCase()+parts2[0].slice(1)) : '';
    s.LastName  = parts2[1] ? (parts2[1].charAt(0).toUpperCase()+parts2[1].slice(1)) : '';
  }
  // Ensure StudentID exists
  if (!s.StudentID && s.studentId) s.StudentID = s.studentId;
  // Ensure Status has a default
  if (!s.Status) s.Status = 'Active';
  return s;
}

function normalizeTeacherRecord(t) {
  if (!t) return t;
  // Clean undefined/null strings
  Object.keys(t).forEach(function(k){
    if (typeof t[k] === 'string') {
      t[k] = t[k].trim();
      if (t[k] === 'undefined' || t[k] === 'null') t[k] = '';
    }
  });
  if (!t.FirstName && !t.LastName && t.Name) {
    var parts = String(t.Name).trim().split(/\s+/);
    t.FirstName = parts[0] || '';
    t.LastName  = parts.slice(1).join(' ') || '';
  }
  if (!t.FirstName && t.Email) {
    var ep = String(t.Email).split('@')[0].replace(/[^a-zA-Z]/g,' ').trim();
    var parts2 = ep.split(/\s+/);
    t.FirstName = parts2[0] ? (parts2[0].charAt(0).toUpperCase()+parts2[0].slice(1)) : '';
    t.LastName  = '';
  }
  if (!t.TeacherID && t.teacherId) t.TeacherID = t.teacherId;
  if (!t.Status) t.Status = 'Active';
  return t;
}

function parseSheet(ss, shName) {
  try {
    var sh = ss.getSheetByName(shName);
    if (!sh) return [];
    var lr = sh.getLastRow();
    if (lr < 2) return [];
    var lc = sh.getLastColumn();
    if (lc < 1) return [];
    var vals = sh.getRange(1, 1, lr, lc).getValues();
    var rawHeaders = vals[0];
    // Normalize header keys so "First Name", "first_name", "FirstName" all → "FirstName"
    var headers = rawHeaders.map(normalizeKey);
    var tz = Session.getScriptTimeZone();
    var out = [];
    for (var i = 1; i < vals.length; i++) {
      var row = vals[i];
      var hasData = false;
      for (var c = 0; c < row.length; c++) {
        if (row[c] !== '' && row[c] !== null) { hasData = true; break; }
      }
      if (!hasData) continue;
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        if (!headers[j]) continue;
        var v = row[j];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        obj[headers[j]] = (v !== undefined && v !== null) ? v : '';
      }
      out.push(obj);
    }
    return out;
  } catch(e) { return []; }
}

// ─── getAllData: one call, row-limited to stay under 6MB ───
function getAllData() {
  try {
    var ss = getSS();
    function safe(fn){ try { return fn(); } catch(e){ return []; } }
    // Use CacheService to avoid re-reading static sheets on every call
    var cache = CacheService.getScriptCache();

    // Settings — build a key→value map (cached 5 min)
    var settings = {};
    var settCache = cache.get('alldata_settings');
    if(settCache){ try{settings=JSON.parse(settCache);}catch(e){} }
    if(!Object.keys(settings).length){
      safe(function(){ return parseSheet(ss, SH.SETTINGS); })
        .forEach(function(s){ if(s.Key) settings[s.Key] = s.Value; });
      try{cache.put('alldata_settings',JSON.stringify(settings),300);}catch(e){}
    }

    return {
      settings:     settings,
      students:     safe(function(){ return parseSheet(ss, SH.STUDENTS).map(normalizeStudentRecord); }),
      teachers:     safe(function(){ return parseSheet(ss, SH.TEACHERS).map(normalizeTeacherRecord); }),
      // Cache semi-static sheets for 3 minutes
      classes:      safe(function(){ return getCachedSheet(ss,cache,SH.CLASSES,180); }),
      sections:     safe(function(){ return getCachedSheet(ss,cache,SH.SECTIONS,180); }),
      subjects:     safe(function(){ return getCachedSheet(ss,cache,SH.SUBJECTS,180); }),
      exams:        safe(function(){ return parseSheet(ss, SH.EXAMS); }), // exams change publish status
      sessions:     safe(function(){ return getCachedSheet(ss,cache,SH.SESSIONS,300); }),
      grading:      safe(function(){ return getCachedSheet(ss,cache,SH.GRADING,300); }),
      resultPublish:safe(function(){ return parseSheet(ss, SH.RESULT_PUBLISH); }),
      feeStructure: safe(function(){ return getCachedSheet(ss,cache,SH.FEE_STRUCTURE,30); }),
      timetable:    safe(function(){ return getCachedSheet(ss,cache,SH.TIMETABLE,300); }),
      marks:        safe(function(){ return parseSheetLimited(ss, SH.MARKS,       2000); }),
      results:      safe(function(){ return parseSheetLimited(ss, SH.RESULTS,     1000); }),
      attendance:   safe(function(){ return parseSheetLimited(ss, SH.ATTENDANCE,  1500); }),
      fees:         safe(function(){ return parseSheetLimited(ss, SH.FEES,        1000); }),
      notifications:safe(function(){ return parseSheetLimited(ss, SH.NOTIFICATIONS,200); }),
      loadedAt: new Date().toISOString()
    };
  } catch(e) { return {error: e.message}; }
}

// Load last N rows of a sheet (avoids 6MB GAS return limit)
function parseSheetLimited(ss, shName, maxRows) {
  try {
    var sh = ss.getSheetByName(shName);
    if (!sh) return [];
    var lr = sh.getLastRow();
    if (lr < 2) return [];
    var lc = sh.getLastColumn();
    if (lc < 1) return [];
    var tz = Session.getScriptTimeZone();
    var rawHdrs = sh.getRange(1, 1, 1, lc).getValues()[0];
    var headers = rawHdrs.map(normalizeKey);
    // Read only last maxRows rows
    var startRow = Math.max(2, lr - maxRows + 1);
    var numRows = lr - startRow + 1;
    var vals = sh.getRange(startRow, 1, numRows, lc).getValues();
    var out = [];
    for (var i = 0; i < vals.length; i++) {
      var row = vals[i];
      var hasData = false;
      for (var c = 0; c < row.length; c++) {
        if (row[c] !== '' && row[c] !== null) { hasData = true; break; }
      }
      if (!hasData) continue;
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var v = row[j];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        obj[headers[j]] = (v !== undefined && v !== null) ? v : '';
      }
      out.push(obj);
    }
    return out;
  } catch(e) { return []; }
}

function deleteStudent(id){
  var old=sheetToObjects(SH.STUDENTS).find(function(s){return String(s.StudentID)===String(id);});
  var res=removeRow(SH.STUDENTS,'StudentID',id);
  if(res.success) addAudit('DELETE',id,SH.STUDENTS,'system','admin',old,null,'Student deleted');
  return res;
}

function saveStudent(d){
  try {
    // Sanitize inputs — no recursion, just string cleaning
    d = sanitizeStudent(d);
    var isNew = !d.StudentID;

    if(isNew){
      d.StudentID = 'STU_' + Date.now();
      if(!d.RollNo || d.RollNo === 'AUTO'){
        try { d.RollNo = generateAdmissionNumber(d.AdmissionYear || new Date().getFullYear()); }
        catch(e){ d.RollNo = String(new Date().getFullYear()) + '-' + Date.now().toString().slice(-4); }
      }
      d.CreatedAt = new Date().toISOString();
    }

    // Write to sheet
    var res = upsertRow(SH.STUDENTS, 'StudentID', d);
    if(!res || !res.success) return res;

    // Audit (safe — truncate to avoid JSON.stringify blowing stack on large objects)
    try {
      var auditNew = { StudentID:d.StudentID, FirstName:d.FirstName, LastName:d.LastName,
                       Email:d.Email, ClassID:d.ClassID, Status:d.Status };
      addAudit(isNew?'CREATE':'UPDATE', d.StudentID, SH.STUDENTS,
               d.UpdatedBy||'system', 'admin', null, auditNew,
               isNew?'New student added':'Student updated');
      SpreadsheetApp.flush();
    } catch(auditErr){ Logger.log('Audit error: '+auditErr.message); }

    // Create login account for new student (best-effort, never fails the main save)
    if(isNew && d.Email && String(d.Email).indexOf('@') > 0){
      try {
        var ss = getSS();
        var usersSh = ss.getSheetByName(SH.USERS);
        if(usersSh){
          var usersData = usersSh.getDataRange().getValues();
          var emailCol = usersData[0].indexOf('Email');
          var alreadyExists = false;
          if(emailCol >= 0){
            for(var i=1;i<usersData.length;i++){
              if(String(usersData[i][emailCol]).trim().toLowerCase() ===
                 String(d.Email).trim().toLowerCase()){ alreadyExists=true; break; }
            }
          }
          if(!alreadyExists){
            var defaultPass = generateDefaultPassword(d.FirstName, d.RollNo);
            upsertRow(SH.USERS, 'UserID', {
              UserID: 'USR_' + Date.now(),
              Email: String(d.Email).trim().toLowerCase(),
              Password: hashPassword(defaultPass),
              Role: 'student',
              Name: (d.FirstName||'')+' '+(d.LastName||''),
              Avatar: ((d.FirstName||'')[0]||'').toUpperCase(),
              AssociatedID: d.StudentID,
              Status: 'Active',
              LastLogin: '',
              CreatedAt: new Date().toISOString(),
              MustChangePassword: 'TRUE',
              ResetOTP: '',
              OTPExpiry: ''
            });
            SpreadsheetApp.flush();
            try { sendWelcomeEmail(d, defaultPass); } catch(me){}
          }
        }
      } catch(acctErr){ Logger.log('Account create error: '+acctErr.message); }
    }

    return { success: true, action: isNew?'created':'updated', id: d.StudentID };
  } catch(e){
    return { success: false, message: 'saveStudent error: ' + e.message };
  }
}

// ── Generate a readable default password ──────────────────────────
function generateDefaultPassword(firstName, rollNo) {
  var name = String(firstName||'student').replace(/[^a-zA-Z]/g,'').toLowerCase().slice(0,4);
  if(name.length < 2) name = 'pass';
  var roll = String(rollNo||'').replace(/[^0-9]/g,'').slice(-4);
  if(!roll) roll = String(Math.floor(1000+Math.random()*9000));
  return name + '@' + roll;   // e.g.  john@1234
}

// ── Send welcome email with default credentials ────────────────────
function sendWelcomeEmail(student, defaultPass) {
  var settings = getCachedSettings ? getCachedSettings() : {};
  var schoolName = settings['school_name'] || 'School';
  var loginUrl = ScriptApp.getService().getUrl();
  var name = ((student.FirstName||'')+' '+(student.LastName||'')).trim();

  var subject = 'Welcome to ' + schoolName + ' - Your Login Details';

  var line = '----------------------------';
  var body = 'Dear ' + name + ',\n\n'
    + 'Your student account has been created at ' + schoolName + '.\n\n'
    + line + '\n'
    + '  LOGIN CREDENTIALS\n'
    + line + '\n'
    + '  Portal URL : ' + loginUrl + '\n'
    + '  Email      : ' + student.Email + '\n'
    + '  Password   : ' + defaultPass + '\n'
    + line + '\n\n'
    + 'IMPORTANT: You must change this password on your first login.\n'
    + 'Please choose a strong personal password that you do not share.\n\n'
    + 'If you did not expect this email, please contact the school.\n\n'
    + 'Regards,\n' + schoolName;

  var recipients = [];
  if (student.Email && student.Email.indexOf('@') > 0)
    recipients.push(student.Email);
  if (student.GuardianEmail && student.GuardianEmail.indexOf('@') > 0
      && student.GuardianEmail !== student.Email)
    recipients.push(student.GuardianEmail);

  recipients.forEach(function(addr) {
    try { MailApp.sendEmail({ to: addr, subject: subject, body: body }); } catch(e) {}
  });
}
// Auto-generate unique admission number: YEAR-XXXX
function generateAdmissionNumber(admYear) {
  var year = admYear || new Date().getFullYear();
  // Get settings for prefix
  var settings = {};
  sheetToObjects(SH.SETTINGS).forEach(function(s){ settings[s.Key] = s.Value; });
  var prefix = settings['admission_prefix'] || '';
  // Find last admission number for this year
  var students = sheetToObjects(SH.STUDENTS);
  var yearStr = String(year);
  var maxSeq = 0;
  students.forEach(function(s) {
    var rn = String(s.RollNo || '');
    // Match format: [PREFIX-]YEAR-NNNN
    var parts = rn.split('-');
    var yearPart = prefix ? parts[1] : parts[0];
    var seqPart = prefix ? parts[2] : parts[1];
    if (yearPart === yearStr && seqPart) {
      var seq = parseInt(seqPart, 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });
  var nextSeq = String(maxSeq + 1).padStart(4, '0');
  return prefix ? (prefix + '-' + yearStr + '-' + nextSeq) : (yearStr + '-' + nextSeq);
}
function updateStudentPhoto(studentId,photoData){
  try {
    if(!photoData||!photoData.trim()) return {success:false,message:'No photo data provided.'};
    if(photoData.length>300000) return {success:false,message:'Image too large (max ~200KB). Please compress and retry.'};
    var ss=getSS();var sh=ss.getSheetByName(SH.STUDENTS);
    var lr=sh.getLastRow();if(lr<2)return {success:false};
    var data=sh.getRange(1,1,lr,sh.getLastColumn()).getValues();
    var headers=data[0];var si=headers.indexOf('StudentID'),pi=headers.indexOf('PhotoURL');
    if(pi<0){ sh.getRange(1,headers.length+1).setValue('PhotoURL'); pi=headers.length; }
    for(var i=1;i<data.length;i++){
      if(String(data[i][si])===String(studentId)){
        sh.getRange(i+1,pi+1).setValue(photoData);
        SpreadsheetApp.flush(); clearSheetCache(SH.STUDENTS);
        addAudit('UPDATE',studentId,SH.STUDENTS,'system','admin',null,{photo:'updated'},'Student photo updated');
        SpreadsheetApp.flush(); return {success:true};
      }
    }
    return {success:false,message:'Student not found'};
  }catch(e){return {success:false,message:e.message};}
}

function updateTeacherPhoto(teacherId, photoData) {
  try {
    if(!photoData||!photoData.trim()) return {success:false,message:'No photo data provided.'};
    if(photoData.length>300000) return {success:false,message:'Image too large (max ~200KB).'};
    var ss=getSS(); var sh=ss.getSheetByName(SH.TEACHERS);
    var lr=sh.getLastRow(); if(lr<2) return {success:false};
    var data=sh.getRange(1,1,lr,sh.getLastColumn()).getValues();
    var hdr=data[0]; var ti=hdr.indexOf('TeacherID'), pi=hdr.indexOf('PhotoURL');
    if(pi<0){ sh.getRange(1,hdr.length+1).setValue('PhotoURL'); pi=hdr.length; }
    for(var i=1;i<data.length;i++){
      if(String(data[i][ti])===String(teacherId)){
        sh.getRange(i+1,pi+1).setValue(photoData);
        SpreadsheetApp.flush(); clearSheetCache(SH.TEACHERS);
        addAudit('UPDATE',teacherId,SH.TEACHERS,'system','admin',null,{photo:'updated'},'Teacher photo updated');
        SpreadsheetApp.flush(); return {success:true};
      }
    }
    return {success:false,message:'Teacher not found'};
  } catch(e){ return {success:false,message:e.message}; }
}

// ═══════════════════════════════════════════════════════════
//  TEACHERS
// ═══════════════════════════════════════════════════════════
function saveTeacher(d){
  try {
    d = sanitizeTeacher(d);
    var isNew = !d.TeacherID;
    if(isNew){
      var teacherCount = 0;
      try { teacherCount = getSS().getSheetByName(SH.TEACHERS).getLastRow() - 1; } catch(e){}
      d.TeacherID = 'TCH_' + Date.now();
      d.EmpID = 'EMP-' + new Date().getFullYear() + '-' + String(Math.max(teacherCount+1,1)).padStart(3,'0');
      d.CreatedAt = new Date().toISOString();
    }
    // Ensure AssignedSubjects and ClassTeacherOf are persisted
    if(d.AssignedSubjects===undefined) d.AssignedSubjects='';
    if(d.ClassTeacherOf===undefined) d.ClassTeacherOf='';
    var res = upsertRow(SH.TEACHERS, 'TeacherID', d);
    if(!res || !res.success) return res;

    // Create login account for new teacher
    if(isNew && d.Email && String(d.Email).indexOf('@') > 0){
      try {
        var ss = getSS();
        var usersSh = ss.getSheetByName(SH.USERS);
        if(usersSh){
          var usersData = usersSh.getDataRange().getValues();
          var emailCol = usersData[0].indexOf('Email');
          var alreadyExists = false;
          if(emailCol >= 0){
            for(var i=1;i<usersData.length;i++){
              if(String(usersData[i][emailCol]).trim().toLowerCase() ===
                 String(d.Email).trim().toLowerCase()){ alreadyExists=true; break; }
            }
          }
          if(!alreadyExists){
            var defaultPass = generateDefaultPassword(d.FirstName || 'teacher', d.EmpID);
            upsertRow(SH.USERS, 'UserID', {
              UserID: 'USR_' + Date.now(),
              Email: String(d.Email).trim().toLowerCase(),
              Password: hashPassword(defaultPass),
              Role: 'teacher',
              Name: (d.FirstName||'')+' '+(d.LastName||''),
              Avatar: ((d.FirstName||'')[0]||'').toUpperCase(),
              AssociatedID: d.TeacherID,
              Status: 'Active',
              LastLogin: '',
              CreatedAt: new Date().toISOString(),
              MustChangePassword: 'TRUE',
              ResetOTP: '',
              OTPExpiry: ''
            });
            SpreadsheetApp.flush();
            try { sendWelcomeEmail(d, defaultPass); } catch(me){}
          }
        }
      } catch(acctErr){ Logger.log('Teacher account error: '+acctErr.message); }
    }

    try { addAudit(isNew?'CREATE':'UPDATE', d.TeacherID, SH.TEACHERS, d.UpdatedBy||'system', 'admin',
      null, {TeacherID:d.TeacherID,FirstName:d.FirstName,LastName:d.LastName,Email:d.Email},
      isNew?'New teacher added':'Teacher updated'); SpreadsheetApp.flush(); } catch(e){}

    return { success: true, action: isNew?'created':'updated', id: d.TeacherID, teacherId: d.TeacherID };
  } catch(e){
    return { success: false, message: 'saveTeacher error: ' + e.message };
  }
}

function sanitizeTeacher(d) {
  var strFields = ['FirstName','LastName','Email','Phone','Qualification','Specialization','Address'];
  strFields.forEach(function(f){ if(d[f] !== undefined) d[f] = sanitizeStr(d[f]); });
  return d;
}
function deleteTeacher(id){return removeRow(SH.TEACHERS,'TeacherID',id);}

// ═══════════════════════════════════════════════════════════
//  CLASSES & SECTIONS
// ═══════════════════════════════════════════════════════════
function saveClass(d){if(!d.ClassID)d.ClassID='CLS_'+Date.now();return upsertRow(SH.CLASSES,'ClassID',d);}
function saveSection(d){if(!d.SectionID)d.SectionID='SEC_'+Date.now();return upsertRow(SH.SECTIONS,'SectionID',d);}
function deleteClass(id){return removeRow(SH.CLASSES,'ClassID',id);}
function deleteSection(id){return removeRow(SH.SECTIONS,'SectionID',id);}

// ═══════════════════════════════════════════════════════════
//  SUBJECTS & EXAMS
// ═══════════════════════════════════════════════════════════
function saveSubject(d){if(!d.SubjectID)d.SubjectID='SUBJ_'+Date.now();return upsertRow(SH.SUBJECTS,'SubjectID',d);}
function deleteSubject(id){return removeRow(SH.SUBJECTS,'SubjectID',id);}
function saveExam(d){
  var isNew = !d.ExamID;
  if(isNew){
    d.ExamID = 'EXM_'+Date.now();
    d.CreatedAt = new Date().toISOString();
    // Honour auto-publish setting
    if(!d.hasOwnProperty('Published')){
      var settings = getCachedSettings();
      d.Published = (settings['auto_publish_exams']==='true') ? 'TRUE' : 'FALSE';
    }
  }
  if(!d.Published) d.Published = 'FALSE';
  var res = upsertRow(SH.EXAMS,'ExamID',d);
  if(res.success) addAudit(isNew?'CREATE':'UPDATE',d.ExamID,SH.EXAMS,'system','admin',null,d,(isNew?'Exam created':'Exam updated'));
  return res;
}

// Toggle exam publish status
function toggleExamPublish(examId, publish, adminId) {
  try {
    var ss = getSS(); var sh = ss.getSheetByName(SH.EXAMS);
    var data = sh.getDataRange().getValues(); var hdr = data[0];
    var eidCol = hdr.indexOf('ExamID'); var pubCol = hdr.indexOf('Published');
    if(pubCol < 0) { pubCol = hdr.length; sh.getRange(1,pubCol+1).setValue('Published'); }
    for(var i=1;i<data.length;i++){
      if(String(data[i][eidCol])===String(examId)){
        sh.getRange(i+1,pubCol+1).setValue(publish?'TRUE':'FALSE');
        SpreadsheetApp.flush(); clearSheetCache(SH.EXAMS);
        addAudit(publish?'EXAM_PUBLISH':'EXAM_UNPUBLISH',examId,SH.EXAMS,adminId,'admin',null,null,(publish?'Exam published':'Exam unpublished'));
        SpreadsheetApp.flush();
        return {success:true};
      }
    }
    return {success:false,message:'Exam not found'};
  } catch(e){ return {success:false,message:e.message}; }
}
function deleteExam(id){return removeRow(SH.EXAMS,'ExamID',id);}
function toggleExamPublish(examId,published){
  var ss=getSS();var sh=ss.getSheetByName(SH.EXAMS);
  var lr=sh.getLastRow();if(lr<2)return {success:false};
  var data=sh.getRange(1,1,lr,sh.getLastColumn()).getValues();var headers=data[0];
  var ei=headers.indexOf('ExamID'),pi=headers.indexOf('Published');
  for(var i=1;i<data.length;i++){
    if(String(data[i][ei])===String(examId)){sh.getRange(i+1,pi+1).setValue(published?'TRUE':'FALSE');SpreadsheetApp.flush();return {success:true};}
  }
  return {success:false};
}

// ═══════════════════════════════════════════════════════════
//  MARKS
// ═══════════════════════════════════════════════════════════
// gradingCache param: pass pre-loaded grading array to avoid repeated sheet reads
function computeGrade(pct, gradingCache){
  var g = gradingCache || sheetToObjects(SH.GRADING);
  // Sort descending so we match the highest band first (handles boundary edge cases)
  var sorted = g.slice().sort(function(a,b){ return parseFloat(b.MinPercent)-parseFloat(a.MinPercent); });
  for(var i=0;i<sorted.length;i++){
    if(pct>=parseFloat(sorted[i].MinPercent)) return sorted[i].GradeName;
  }
  return 'F';
}

function saveMarksForSubject(examId, classId, subjectId, teacherId, marksArray){
  try {
    var subjects = sheetToObjects(SH.SUBJECTS);
    var sub = subjects.find(function(s){return String(s.SubjectID)===String(subjectId);});
    if(!sub) return {success:false, message:'Subject not found: '+subjectId};

    // Permission check
    var adminUsers = (sheetToObjects(SH.USERS)||[]).filter(function(u){return String(u.Role||'').toLowerCase()==='admin';});
    var adminIds = adminUsers.map(function(u){return String(u.AssociatedID||'');});
    var isAdmin = !teacherId || adminIds.indexOf(String(teacherId)) >= 0;
    if(!isAdmin && String(sub.TeacherID) !== String(teacherId)){
      var tch = sheetToObjects(SH.TEACHERS).find(function(t){return String(t.TeacherID)===String(teacherId);});
      if(tch){
        var asgSubs = (tch.AssignedSubjects||'').split(',').map(function(s){return s.trim();});
        var asgCls  = (tch.AssignedClasses||'').split(',').map(function(s){return s.trim();});
        if(asgSubs.indexOf(subjectId)<0 && asgCls.indexOf(String(sub.ClassID||classId))<0){
          return {success:false, message:'You are not assigned to this subject. Ask admin to update your profile.'};
        }
      }
    }

    var ss = getSS();
    var sh = ss.getSheetByName(SH.MARKS);
    var lr = sh.getLastRow();
    var maxMarks = parseFloat(sub.MaxMarks||100);
    var gradingCache = sheetToObjects(SH.GRADING);
    var now = new Date().toISOString();

    // Build a map of existing marks for fast lookup (studentId → rowIndex)
    var existingMap = {}; // studentId → {rowNum, rowData}
    if(lr >= 2){
      var allData = sh.getRange(1,1,lr,sh.getLastColumn()).getValues();
      var hdr = allData[0];
      var eiI=hdr.indexOf('ExamID'), ciI=hdr.indexOf('ClassID'), siI=hdr.indexOf('SubjectID'), stuI=hdr.indexOf('StudentID');
      for(var i=1;i<allData.length;i++){
        if(String(allData[i][eiI])===String(examId) &&
           String(allData[i][ciI])===String(classId) &&
           String(allData[i][siI])===String(subjectId)){
          existingMap[String(allData[i][stuI])] = {rowNum:i+1, hdr:hdr};
        }
      }
    }

    var newRows = []; var updatedCount = 0;
    marksArray.forEach(function(m, idx){
      if(!m.studentId || m.marks==='' || m.marks===null || m.marks===undefined) return;
      var marks = parseFloat(m.marks);
      if(isNaN(marks)||marks<0||marks>maxMarks) return;
      var pct = maxMarks>0 ? marks/maxMarks*100 : 0;
      var grade = computeGrade(pct, gradingCache);

      if(existingMap[String(m.studentId)]){
        // UPDATE existing row in place
        var info = existingMap[String(m.studentId)];
        var h = info.hdr;
        var row = sh.getRange(info.rowNum, 1, 1, h.length).getValues()[0];
        var moI=h.indexOf('MarksObtained');
        var grI=h.indexOf('Grade');
        var upI=h.indexOf('UpdatedAt');
        if(moI>=0) row[moI]=marks;
        if(grI>=0) row[grI]=grade;
        if(upI>=0) row[upI]=now;
        sh.getRange(info.rowNum,1,1,h.length).setValues([row]);
        updatedCount++;
      } else {
        // INSERT new row
        newRows.push([
          'MRK_'+Date.now()+'_'+idx+'_'+Math.random().toString(36).slice(2,5),
          examId, m.studentId, subjectId, classId, teacherId||'',
          marks, maxMarks, grade, '', now, now
        ]);
      }
    });

    if(newRows.length>0){
      // Ensure headers exist
      if(sh.getLastRow()===0){
        sh.appendRow(['MarkID','ExamID','StudentID','SubjectID','ClassID','TeacherID','MarksObtained','MaxMarks','Grade','Remarks','CreatedAt','UpdatedAt']);
      }
      sh.getRange(sh.getLastRow()+1,1,newRows.length,newRows[0].length).setValues(newRows);
    }
    SpreadsheetApp.flush();
    clearSheetCache(SH.MARKS);
    addAudit('MARKS_UPLOAD',examId+'/'+subjectId,SH.MARKS,teacherId,'teacher',null,{count:marksArray.length},'Marks saved for '+subjectId);
    SpreadsheetApp.flush();
    return {success:true, count:newRows.length+updatedCount};
  } catch(e){ return {success:false, message:'saveMarksForSubject: '+e.message}; }
}

function generateResults(examId){
  try {
    var marks=sheetToObjects(SH.MARKS).filter(function(m){return String(m.ExamID)===String(examId);});
    var sids=[];marks.forEach(function(m){if(sids.indexOf(m.StudentID)<0)sids.push(m.StudentID);});
    var ss=getSS(); var sh=ss.getSheetByName(SH.RESULTS);
    var lr=sh.getLastRow();
    // Batch delete existing results for this exam
    if(lr>=2){
      var data=sh.getRange(1,1,lr,sh.getLastColumn()).getValues();
      var ei=data[0].indexOf('ExamID');
      for(var i=lr;i>=2;i--){if(String(data[i-1][ei])===String(examId))sh.deleteRow(i);}
    }
    // Cache grading once
    var gradingCache = sheetToObjects(SH.GRADING);
    var calcs=sids.map(function(sid){
      var sm=marks.filter(function(m){return String(m.StudentID)===String(sid);});
      var obt=sm.reduce(function(s,m){return s+parseFloat(m.MarksObtained||0);},0);
      var mx=sm.reduce(function(s,m){return s+parseFloat(m.MaxMarks||0);},0);
      var p=mx>0?parseFloat((obt/mx*100).toFixed(2)):0;
      return {sid:sid,obt:obt,mx:mx,p:p};
    }).sort(function(a,b){return b.p-a.p;});
    // Assign ranks with proper tie handling
    var rank=1;
    calcs.forEach(function(r,i){
      if(i>0&&r.p===calcs[i-1].p){ r.rank=calcs[i-1].rank; }
      else{ r.rank=rank; }
      rank++;
    });
    // Batch write all results in one setValues call
    var now=new Date().toISOString();
    var newRows=calcs.map(function(r,i){
      return ['RES_'+Date.now()+'_'+i,examId,r.sid,r.mx,r.obt,r.p.toFixed(2),
        computeGrade(r.p,gradingCache),r.rank,'FALSE',now];
    });
    if(newRows.length>0){
      var startRow=sh.getLastRow()+1;
      sh.getRange(startRow,1,newRows.length,newRows[0].length).setValues(newRows);
    }
    SpreadsheetApp.flush();
    clearSheetCache(SH.RESULTS);
    addAudit('GENERATE_RESULTS',examId,SH.RESULTS,'system','admin',null,{count:calcs.length},'Results generated');
    SpreadsheetApp.flush();
    return {success:true,count:calcs.length};
  }catch(e){return {success:false,message:e.message};}
}

// ═══════════════════════════════════════════════════════════
//  RESULT PUBLISH SYSTEM
// ═══════════════════════════════════════════════════════════
function publishResult(examId,classId,examType,publishStatus,publishedBy){
  try {
    var all=sheetToObjects(SH.RESULT_PUBLISH);
    var existing=null;
    for(var i=0;i<all.length;i++){if(String(all[i].ExamID)===String(examId)&&String(all[i].ClassID)===String(classId)){existing=all[i];break;}}
    var data={ExamID:examId,ClassID:classId,ExamType:examType,PublishStatus:publishStatus?'TRUE':'FALSE',PublishedBy:publishedBy,PublishedAt:publishStatus?new Date().toISOString():''};
    if(existing)data.PublishID=existing.PublishID;else data.PublishID='PUB_'+Date.now();
    upsertRow(SH.RESULT_PUBLISH,'PublishID',data);
    toggleExamPublish(examId,publishStatus);
    // Update results sheet — filter by BOTH examId AND classId to avoid cross-class pollution
    var ss=getSS(); var sh=ss.getSheetByName(SH.RESULTS);
    var lr=sh.getLastRow();if(lr>=2){var d=sh.getRange(1,1,lr,sh.getLastColumn()).getValues();var h=d[0];var ei=h.indexOf('ExamID'),pi=h.indexOf('Published'),ci=h.indexOf('StudentID');
      // Build studentId set for this class
      var classStuIds={};sheetToObjects(SH.STUDENTS).filter(function(s){return String(s.ClassID)===String(classId);}).forEach(function(s){classStuIds[s.StudentID]=true;});
      for(var i=1;i<d.length;i++){if(String(d[i][ei])===String(examId)&&classStuIds[d[i][ci]])sh.getRange(i+1,pi+1).setValue(publishStatus?'TRUE':'FALSE');}
    }
    SpreadsheetApp.flush();
    addAudit('PUBLISH_RESULT',examId,SH.RESULT_PUBLISH,publishedBy,'admin',{status:!publishStatus},{status:publishStatus},(publishStatus?'Results published':'Results unpublished'));
    return {success:true};
  }catch(e){return {success:false,message:e.message};}
}

// ═══════════════════════════════════════════════════════════
//  STUDENT RESULTS (published only)
// ═══════════════════════════════════════════════════════════
function getStudentResults(studentId){
  var ss=getSS();
  var allStudents=parseSheet(ss,SH.STUDENTS);
  var student=allStudents.find(function(s){return String(s.StudentID)===String(studentId);});
  if(!student) return {success:false,message:'Student not found'};
  var allMarks=parseSheet(ss,SH.MARKS).filter(function(m){return String(m.StudentID)===String(studentId);});
  var allExams=parseSheet(ss,SH.EXAMS);
  var allSubjects=parseSheet(ss,SH.SUBJECTS);
  var allResults=parseSheet(ss,SH.RESULTS);
  var allPub=parseSheet(ss,SH.RESULT_PUBLISH);
  var gradingCache=parseSheet(ss,SH.GRADING);
  var settings={};parseSheet(ss,SH.SETTINGS).forEach(function(s){settings[s.Key]=s.Value;});
  // Build exam marks index for O(1) lookup
  var marksByExam={};
  allMarks.forEach(function(m){
    if(!marksByExam[m.ExamID]) marksByExam[m.ExamID]=[];
    marksByExam[m.ExamID].push(m);
  });
  var publishedExams=[];
  allExams.forEach(function(ex){
    var pub=allPub.find(function(p){return String(p.ExamID)===String(ex.ExamID)&&String(p.ClassID)===String(student.ClassID);});
    if(!pub||String(pub.PublishStatus)!=='TRUE')return;
    var examMarks=marksByExam[ex.ExamID]||[];
    if(!examMarks.length)return;
    var subjects=allSubjects.filter(function(s){return String(s.ClassID)===String(student.ClassID);});
    var subjectDetails=subjects.map(function(sub){
      var m=examMarks.find(function(mk){return String(mk.SubjectID)===String(sub.SubjectID);});
      var obt=m?parseFloat(m.MarksObtained):0;
      var mx=m?parseFloat(m.MaxMarks):parseFloat(sub.MaxMarks)||100;
      var pct=mx>0?(obt/mx*100):0;
      // Use subject-specific PassMarks instead of hardcoded 40%
      var passMarks=parseFloat(sub.PassMarks)||0;
      var passPct=mx>0?(passMarks/mx*100):40;
      return {subjectId:sub.SubjectID,name:sub.SubjectName,code:sub.SubjectCode,
        obtained:obt,max:mx,percentage:pct.toFixed(1),
        grade:m?m.Grade:computeGrade(pct,gradingCache),pass:pct>=passPct};
    });
    var totObt=subjectDetails.reduce(function(s,d){return s+d.obtained;},0);
    var totMax=subjectDetails.reduce(function(s,d){return s+d.max;},0);
    var overall=totMax>0?(totObt/totMax*100).toFixed(2):0;
    var result=allResults.find(function(r){return String(r.ExamID)===String(ex.ExamID)&&String(r.StudentID)===String(studentId);});
    publishedExams.push({exam:ex,subjectDetails:subjectDetails,totalObtained:totObt,totalMax:totMax,
      percentage:overall,grade:computeGrade(parseFloat(overall),gradingCache),rank:result?result.Rank:'N/A'});
  });
  return {success:true,student:student,results:publishedExams,settings:settings};
}

function getClassResults(examId,classId){
  var ss=getSS();
  var students=parseSheet(ss,SH.STUDENTS).filter(function(s){return String(s.ClassID)===String(classId);});
  var subjects=parseSheet(ss,SH.SUBJECTS).filter(function(s){return String(s.ClassID)===String(classId);});
  var allMarks=parseSheet(ss,SH.MARKS).filter(function(m){return String(m.ExamID)===String(examId)&&String(m.ClassID)===String(classId);});
  var results=parseSheet(ss,SH.RESULTS).filter(function(r){return String(r.ExamID)===String(examId);});
  var exam=parseSheet(ss,SH.EXAMS).find(function(e){return String(e.ExamID)===String(examId);});
  var gradingCache=parseSheet(ss,SH.GRADING);
  var settings={};parseSheet(ss,SH.SETTINGS).forEach(function(s){settings[s.Key]=s.Value;});
  // Build O(1) index maps
  var marksByStu={};
  allMarks.forEach(function(m){
    if(!marksByStu[m.StudentID]) marksByStu[m.StudentID]={};
    marksByStu[m.StudentID][m.SubjectID]=m;
  });
  var resultsIdx={};
  results.forEach(function(r){resultsIdx[r.StudentID]=r;});
  var studentData=students.map(function(stu){
    var stuMarks=marksByStu[stu.StudentID]||{};
    var subDetails=subjects.map(function(sub){
      var m=stuMarks[sub.SubjectID];
      var obt=m?parseFloat(m.MarksObtained):0;
      var mx=m?parseFloat(m.MaxMarks):parseFloat(sub.MaxMarks)||100;
      var pct=mx>0?(obt/mx*100):0;
      // Use subject-specific PassMarks, not hardcoded 40%
      var passPct=mx>0?(parseFloat(sub.PassMarks||0)/mx*100):40;
      return {name:sub.SubjectName,code:sub.SubjectCode,obtained:obt,max:mx,
        percentage:pct.toFixed(1),grade:m?m.Grade:computeGrade(pct,gradingCache),pass:pct>=passPct};
    });
    var totObt=subDetails.reduce(function(s,d){return s+d.obtained;},0);
    var totMax=subDetails.reduce(function(s,d){return s+d.max;},0);
    var overall=totMax>0?(totObt/totMax*100).toFixed(2):0;
    var result=resultsIdx[stu.StudentID];
    return {student:stu,subjectDetails:subDetails,totalObtained:totObt,totalMax:totMax,
      percentage:overall,grade:computeGrade(parseFloat(overall),gradingCache),rank:result?result.Rank:'N/A'};
  }).sort(function(a,b){return parseFloat(b.percentage)-parseFloat(a.percentage);});
  return {success:true,exam:exam,students:studentData,subjects:subjects,settings:settings};
}

// ═══════════════════════════════════════════════════════════
//  FEE MANAGEMENT v2 — Dynamic Calculation, No Stored Totals
//  Architecture: Students | FeeStructure | Payments
//  CORE RULE: Never store totalPaid or pending — always calculate dynamically
// ═══════════════════════════════════════════════════════════

// ── SHEET SETUP ──────────────────────────────────────────────────────────
function setupFeeSheets() {
  var ss = getSS();
  // FeeStructure: fee templates per class
  mkSheet(ss, SH.FEE_STRUCTURE, [
    'class','monthlyFee','transportFee','otherFee','lateFeePerDay','lateFeeGraceDays','updatedAt'
  ], '#4CAF50');
  // Payments: every transaction — append-only, never overwrite
  mkSheet(ss, SH.PAYMENTS, [
    'receiptNo','studentId','amountPaid','paymentDate','paymentMode','month','notes','collectedBy','createdAt'
  ], '#2196F3');
  SpreadsheetApp.flush();
  return { success: true, message: 'Fee sheets created. Run this once after deployment.' };
}

// ── 1. GET MONTHLY FEE FOR A CLASS ───────────────────────────────────────
function getMonthlyFee(studentClass) {
  var ss = getSS();
  var data = ss.getSheetByName(SH.FEE_STRUCTURE).getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(studentClass).trim()) {
      var monthly = Number(data[i][1]) || 0;
      var transport = Number(data[i][2]) || 0;
      var other = Number(data[i][3]) || 0;
      return {
        success: true,
        monthlyFee: monthly,
        transportFee: transport,
        otherFee: other,
        total: monthly + transport + other,
        lateFeePerDay: Number(data[i][4]) || 0,
        lateFeeGraceDays: Number(data[i][5]) || 5
      };
    }
  }
  return { success: false, message: 'No fee structure found for class: ' + studentClass };
}

// ── 2. GET TOTAL PAID — DYNAMIC, NEVER STORED ────────────────────────────
function getTotalPaid(studentId) {
  var ss = getSS();
  var data = ss.getSheetByName(SH.PAYMENTS).getDataRange().getValues();
  var total = 0;
  var byMonth = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(studentId)) {
      var amt = Number(data[i][2]) || 0;
      var month = String(data[i][5] || '');
      total += amt;
      if (month) {
        byMonth[month] = (byMonth[month] || 0) + amt;
      }
    }
  }
  return { total: total, byMonth: byMonth };
}

// ── 3. CALCULATE TOTAL FEE DUE TILL NOW ──────────────────────────────────
// This is where most systems fail — dynamically computed from admission date
function getTotalFeeTillNow(studentClass, admissionDate) {
  if (!admissionDate) return { success: false, message: 'No admission date' };
  var adDate = admissionDate instanceof Date ? admissionDate : new Date(admissionDate);
  if (isNaN(adDate.getTime())) return { success: false, message: 'Invalid admission date' };
  var today = new Date();
  // Count months from admission month to current month (inclusive)
  var months = (today.getFullYear() - adDate.getFullYear()) * 12 +
               (today.getMonth() - adDate.getMonth()) + 1;
  if (months < 1) months = 1;
  var feeInfo = getMonthlyFee(studentClass);
  if (!feeInfo.success) return feeInfo;
  return {
    success: true,
    months: months,
    monthlyTotal: feeInfo.total,
    totalDue: months * feeInfo.total,
    feeBreakdown: feeInfo
  };
}

// ── 4. CALCULATE PENDING AMOUNT — THE CORE ───────────────────────────────
function getPendingAmount(studentId, studentClass, admissionDate) {
  var feeCalc = getTotalFeeTillNow(studentClass, admissionDate);
  if (!feeCalc.success) return { success: false, message: feeCalc.message };
  var paidCalc = getTotalPaid(studentId);
  var pending = feeCalc.totalDue - paidCalc.total;
  return {
    success: true,
    totalBilled: feeCalc.totalDue,
    totalPaid: paidCalc.total,
    pending: pending,
    months: feeCalc.months,
    monthlyFee: feeCalc.monthlyTotal,
    paidByMonth: paidCalc.byMonth,
    status: pending <= 0 ? 'Clear' : (paidCalc.total > 0 ? 'Partial' : 'Pending')
  };
}

// ── 5. LATE FEE CALCULATION ───────────────────────────────────────────────
function calculateLateFee(studentClass, paymentDate) {
  var feeInfo = getMonthlyFee(studentClass);
  if (!feeInfo.success) return 0;
  var graceDays = feeInfo.lateFeeGraceDays || 5;
  var finePerDay = feeInfo.lateFeePerDay || 0;
  if (!finePerDay) return 0;
  var pDate = paymentDate instanceof Date ? paymentDate : new Date();
  // Fee is due on the 1st; grace period = graceDays
  var dueDay = new Date(pDate.getFullYear(), pDate.getMonth(), graceDays + 1);
  if (pDate > dueDay) {
    var daysLate = Math.floor((pDate - dueDay) / (1000 * 60 * 60 * 24));
    return daysLate * finePerDay;
  }
  return 0;
}

// ── 6. RECORD PAYMENT — APPEND ONLY ──────────────────────────────────────
function makePayment(studentId, amountPaid, month, paymentMode, notes, collectedBy) {
  try {
    var ss = getSS();
    var sh = ss.getSheetByName(SH.PAYMENTS);
    if (!sh) return { success: false, message: 'Payments sheet not found. Run setupFeeSheets() first.' };
    // Validate
    if (!studentId) return { success: false, message: 'studentId is required' };
    if (!amountPaid || isNaN(Number(amountPaid)) || Number(amountPaid) <= 0)
      return { success: false, message: 'amountPaid must be a positive number' };
    if (!month) return { success: false, message: 'month is required (e.g. "April 2025")' };
    var receipt = 'REC' + Date.now();
    var now = new Date();
    sh.appendRow([
      receipt,
      String(studentId),
      Number(amountPaid),
      now,
      paymentMode || 'Cash',
      String(month),
      notes || '',
      collectedBy || '',
      now.toISOString()
    ]);
    SpreadsheetApp.flush();
    // Send email receipt
    try { sendEmailReceipt(studentId, amountPaid, receipt, month); } catch(e) {
      Logger.log('Email failed (non-fatal): ' + e.message);
    }
    return { success: true, receiptNo: receipt, message: 'Payment recorded successfully' };
  } catch(e) {
    return { success: false, message: 'makePayment error: ' + e.message };
  }
}

// ── 7. EMAIL RECEIPT ──────────────────────────────────────────────────────
function sendEmailReceipt(studentId, amount, receiptNo, month) {
  var ss = getSS();
  var students = ss.getSheetByName(SH.STUDENTS).getDataRange().getValues();
  var email = '', name = '', cls = '';
  for (var i = 1; i < students.length; i++) {
    if (String(students[i][0]) === String(studentId)) {
      // Columns: studentId, name, class, section, email (adjust if your sheet differs)
      name  = students[i][1] || '';
      cls   = students[i][2] || '';
      email = students[i][4] || '';
      break;
    }
  }
  if (!email) return; // No email, skip silently
  var settings = {};
  try {
    ss.getSheetByName(SH.SETTINGS).getDataRange().getValues().forEach(function(r){ if(r[0]) settings[r[0]]=r[1]; });
  } catch(e){}
  var schoolName = settings['school_name'] || 'School';
  var subject = schoolName + ' — Fee Payment Receipt';
  var body = '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden">'
    + '<div style="background:#1e3a8a;color:#fff;padding:16px 20px"><h2 style="margin:0;font-size:18px">' + schoolName + '</h2><p style="margin:4px 0 0;font-size:12px;opacity:.7">Fee Receipt</p></div>'
    + '<div style="padding:20px">'
    + '<p style="margin:0 0 12px">Dear <b>' + name + '</b>,</p>'
    + '<p style="margin:0 0 16px;color:#64748b">Your payment has been received successfully.</p>'
    + '<table style="width:100%;border-collapse:collapse">'
    + '<tr style="background:#f8fafc"><td style="padding:8px 12px;font-weight:600">Receipt No.</td><td style="padding:8px 12px;font-family:monospace">' + receiptNo + '</td></tr>'
    + '<tr><td style="padding:8px 12px;font-weight:600">Amount Paid</td><td style="padding:8px 12px;font-size:18px;font-weight:700;color:#059669">₹' + Number(amount).toFixed(2) + '</td></tr>'
    + '<tr style="background:#f8fafc"><td style="padding:8px 12px;font-weight:600">Month</td><td style="padding:8px 12px">' + month + '</td></tr>'
    + '<tr><td style="padding:8px 12px;font-weight:600">Class</td><td style="padding:8px 12px">' + cls + '</td></tr>'
    + '<tr style="background:#f8fafc"><td style="padding:8px 12px;font-weight:600">Date</td><td style="padding:8px 12px">' + new Date().toLocaleDateString('en-IN') + '</td></tr>'
    + '</table>'
    + '<p style="margin:16px 0 0;font-size:12px;color:#94a3b8">This is a computer-generated receipt. Please keep it for your records.</p>'
    + '</div></div>';
  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
}

// ── 8. GET STUDENT FEE SUMMARY (for profile page) ────────────────────────
function getStudentFeeSummary(studentId) {
  try {
    var ss = getSS();
    // Get student record
    var students = parseSheet(ss, SH.STUDENTS);
    var stu = students.find(function(s){ return String(s.studentId||s.StudentID) === String(studentId); });
    if (!stu) return { success: false, message: 'Student not found' };
    var cls   = stu.class || stu.Class || stu.ClassID || '';
    var admDt = stu.admissionDate || stu.AdmissionDate || stu.createdAt || null;
    // Pending calc
    var pending = getPendingAmount(studentId, cls, admDt);
    // Payment history
    var payments = [];
    var payData = ss.getSheetByName(SH.PAYMENTS).getDataRange().getValues();
    var payHdr  = payData[0];
    for (var i = 1; i < payData.length; i++) {
      if (String(payData[i][1]) === String(studentId)) {
        var row = {};
        payHdr.forEach(function(h, j){ row[h] = payData[i][j]; });
        payments.push(row);
      }
    }
    payments.sort(function(a,b){ return new Date(b.paymentDate||b.createdAt) - new Date(a.paymentDate||a.createdAt); });
    // Fee structure
    var feeInfo = getMonthlyFee(cls);
    return {
      success: true,
      student: stu,
      className: cls,
      feeStructure: feeInfo.success ? feeInfo : null,
      summary: pending.success ? pending : { totalBilled:0, totalPaid:0, pending:0, status:'Unknown' },
      payments: payments,
      generatedAt: new Date().toISOString()
    };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── 9. GET ALL STUDENTS FEE STATUS (for fee management list) ─────────────
function getAllStudentFeeStatus(classFilter) {
  try {
    var ss = getSS();
    var students = parseSheet(ss, SH.STUDENTS).map(normalizeStudentRecord);
    if (classFilter) {
      students = students.filter(function(s){ return String(s.ClassID||s.class||'') === String(classFilter); });
    }
    var result = students.map(function(stu) {
      var cls   = stu.class || stu.Class || stu.ClassID || '';
      var admDt = stu.admissionDate || stu.AdmissionDate || stu.createdAt || null;
      var pending = getPendingAmount(String(stu.StudentID||stu.studentId||''), cls, admDt);
      return {
        studentId:  String(stu.StudentID || stu.studentId || ''),
        name:       stu.FirstName ? (stu.FirstName + ' ' + (stu.LastName||'')).trim() : (stu.name||''),
        class:      cls,
        rollNo:     stu.RollNo || stu.rollNo || '',
        totalBilled: pending.success ? pending.totalBilled : 0,
        totalPaid:   pending.success ? pending.totalPaid   : 0,
        pending:     pending.success ? pending.pending      : 0,
        status:      pending.success ? pending.status       : 'Unknown',
        months:      pending.success ? pending.months       : 0
      };
    });
    var totalPending = result.reduce(function(s,r){ return s + (r.pending>0?r.pending:0); }, 0);
    var totalCollected = result.reduce(function(s,r){ return s + r.totalPaid; }, 0);
    return { success: true, students: result, totalPending: totalPending, totalCollected: totalCollected };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── 10. SAVE / UPDATE FEE STRUCTURE (admin) ───────────────────────────────
function saveFeeStructure(d) {
  try {
    var ss = getSS();
    var sh = ss.getSheetByName(SH.FEE_STRUCTURE);
    if (!sh) return { success: false, message: 'FeeStructure sheet not found. Run setupFeeSheets() first.' };
    if (!d.class) return { success: false, message: '"class" field required' };
    d.monthlyFee   = parseFloat(d.monthlyFee)  || 0;
    d.transportFee = parseFloat(d.transportFee) || 0;
    d.otherFee     = parseFloat(d.otherFee)     || 0;
    d.lateFeePerDay    = parseFloat(d.lateFeePerDay)    || 0;
    d.lateFeeGraceDays = parseFloat(d.lateFeeGraceDays) || 5;
    d.updatedAt = new Date().toISOString();
    // Find existing row for this class
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(d.class).trim()) {
        sh.getRange(i+1, 1, 1, 7).setValues([[
          d.class, d.monthlyFee, d.transportFee, d.otherFee,
          d.lateFeePerDay, d.lateFeeGraceDays, d.updatedAt
        ]]);
        SpreadsheetApp.flush();
        clearSheetCache(SH.FEE_STRUCTURE);
        return { success: true, action: 'updated' };
      }
    }
    // Not found — append
    sh.appendRow([d.class, d.monthlyFee, d.transportFee, d.otherFee, d.lateFeePerDay, d.lateFeeGraceDays, d.updatedAt]);
    SpreadsheetApp.flush();
    clearSheetCache(SH.FEE_STRUCTURE);
    return { success: true, action: 'created' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── 11. GET ALL FEE STRUCTURES ────────────────────────────────────────────
function getAllFeeStructures() {
  try {
    var ss = getSS();
    var sh = ss.getSheetByName(SH.FEE_STRUCTURE);
    if (!sh || sh.getLastRow() < 2) return { success: true, structures: [] };
    var data = sh.getDataRange().getValues();
    var hdr  = data[0];
    var structures = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var row = {};
      hdr.forEach(function(h,j){ row[h] = data[i][j]; });
      structures.push(row);
    }
    return { success: true, structures: structures };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── 12. GET ALL PAYMENTS (filterable) ────────────────────────────────────
function getAllPayments(filters) {
  try {
    var ss = getSS();
    var sh = ss.getSheetByName(SH.PAYMENTS);
    if (!sh || sh.getLastRow() < 2) return { success: true, payments: [] };
    var data = sh.getDataRange().getValues();
    var hdr  = data[0];
    var payments = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var row = {};
      hdr.forEach(function(h,j){ row[h] = data[i][j]; });
      // Apply filters
      if (filters) {
        if (filters.studentId && String(row.studentId) !== String(filters.studentId)) continue;
        if (filters.month     && String(row.month)     !== String(filters.month))     continue;
        if (filters.mode      && String(row.paymentMode) !== String(filters.mode))    continue;
      }
      payments.push(row);
    }
    // Sort newest first
    payments.sort(function(a,b){ return new Date(b.paymentDate||b.createdAt) - new Date(a.paymentDate||a.createdAt); });
    var total = payments.reduce(function(s,p){ return s + (Number(p.amountPaid)||0); }, 0);
    return { success: true, payments: payments, totalCollected: total };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── 13. GET RECEIPT DATA ──────────────────────────────────────────────────
function getReceiptData(receiptNo) {
  try {
    var ss = getSS();
    var payData = ss.getSheetByName(SH.PAYMENTS).getDataRange().getValues();
    var hdr = payData[0];
    for (var i = 1; i < payData.length; i++) {
      if (String(payData[i][0]) === String(receiptNo)) {
        var p = {};
        hdr.forEach(function(h,j){ p[h] = payData[i][j]; });
        // Get student info
        var stuData = ss.getSheetByName(SH.STUDENTS).getDataRange().getValues();
        var stuHdr  = stuData[0];
        for (var j = 1; j < stuData.length; j++) {
          if (String(stuData[j][0]) === String(p.studentId)) {
            var stu = {};
            stuHdr.forEach(function(h,k){ stu[h] = stuData[j][k]; });
            p.student = stu;
            break;
          }
        }
        // Settings
        var settings = {};
        try{ ss.getSheetByName(SH.SETTINGS).getDataRange().getValues().forEach(function(r){if(r[0])settings[r[0]]=r[1];}); }catch(e){}
        p.settings = settings;
        return { success: true, payment: p };
      }
    }
    return { success: false, message: 'Receipt not found: ' + receiptNo };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── 14. FEE REMINDER (manual trigger) ─────────────────────────────────────
function sendFeeReminders(classFilter) {
  try {
    var status = getAllStudentFeeStatus(classFilter||null);
    if (!status.success) return status;
    var ss = getSS();
    var settings = {};
    try{ ss.getSheetByName(SH.SETTINGS).getDataRange().getValues().forEach(function(r){if(r[0])settings[r[0]]=r[1];}); }catch(e){}
    var schoolName = settings['school_name'] || 'School';
    var sent = 0;
    var students = parseSheet(ss, SH.STUDENTS);
    status.students.filter(function(s){ return s.pending > 0; }).forEach(function(s) {
      var stuRec = students.find(function(r){ return String(r.StudentID||r.studentId||'') === s.studentId; });
      if (!stuRec) return;
      var email = stuRec.Email || stuRec.email || stuRec.GuardianEmail || '';
      if (!email) return;
      try {
        MailApp.sendEmail({
          to: email,
          subject: schoolName + ' — Fee Due Reminder',
          htmlBody: '<p>Dear Parent/Guardian of <b>' + s.name + '</b>,</p>'
            + '<p>A fee payment of <b style="color:#dc2626">₹' + s.pending.toFixed(2) + '</b> is pending for ' + s.months + ' month(s).</p>'
            + '<p>Please pay at the earliest to avoid late charges.</p>'
            + '<p>Thank you,<br>' + schoolName + '</p>'
        });
        sent++;
      } catch(e) {}
    });
    return { success: true, sent: sent, defaulters: status.students.filter(function(s){return s.pending>0;}).length };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── 15. MONTHLY COLLECTION REPORT ────────────────────────────────────────
function getMonthlyCollectionReport(month) {
  try {
    var ss = getSS();
    var payData = ss.getSheetByName(SH.PAYMENTS).getDataRange().getValues();
    var hdr = payData[0];
    var byDate = {}, total = 0, count = 0;
    for (var i = 1; i < payData.length; i++) {
      if (!payData[i][0]) continue;
      var p = {};
      hdr.forEach(function(h,j){ p[h] = payData[i][j]; });
      if (month && String(p.month) !== String(month)) continue;
      var dateKey = p.paymentDate ? new Date(p.paymentDate).toLocaleDateString('en-IN') : 'Unknown';
      byDate[dateKey] = (byDate[dateKey]||0) + (Number(p.amountPaid)||0);
      total += Number(p.amountPaid)||0;
      count++;
    }
    return { success: true, month: month||'All', total: total, count: count, byDate: byDate };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── 16. SAVE SETTINGS (school name etc) ───────────────────────────────────
function saveSettings(settings, updatedBy) {
  try {
    var ss = getSS();
    var sh = ss.getSheetByName(SH.SETTINGS);
    if (!sh) { sh = ss.insertSheet(SH.SETTINGS); sh.appendRow(['Key','Value','UpdatedBy','UpdatedAt']); }
    var now = new Date().toISOString();
    var lr = sh.getLastRow();
    var existingMap = {};
    if (lr >= 1) {
      var allData = sh.getRange(1,1,lr,Math.max(sh.getLastColumn(),2)).getValues();
      for (var i = 0; i < allData.length; i++) {
        var k = String(allData[i][0]||'').trim();
        if (k && k !== 'Key') existingMap[k] = i + 1;
      }
    }
    Object.keys(settings).forEach(function(key) {
      var val = settings[key];
      if (typeof val === 'undefined' || val === null) return;
      if (existingMap.hasOwnProperty(key)) {
        sh.getRange(existingMap[key],1,1,4).setValues([[key, val, updatedBy||'admin', now]]);
      } else {
        sh.appendRow([key, val, updatedBy||'admin', now]);
        existingMap[key] = sh.getLastRow();
      }
    });
    SpreadsheetApp.flush();
    try { var c=CacheService.getScriptCache(); c.remove('alldata_settings'); } catch(e){}
    return { success: true };
  } catch(e) { return { success: false, message: 'saveSettings: '+e.message }; }
}

// ── Expose loginSecure as alias (frontend calls loginSecure) ──────────────
var loginSecure = login;

// ── Update getAllData to include fee data from new schema ─────────────────
function getFeeDataForAll() {
  try {
    var ss = getSS();
    function safe(fn){ try{return fn();}catch(e){return [];} }
    var cache = CacheService.getScriptCache();
    return {
      feeStructure: safe(function(){ return getCachedSheet(ss,cache,SH.FEE_STRUCTURE,30); }),
      payments:     safe(function(){ 
        var sh = ss.getSheetByName(SH.PAYMENTS);
        if(!sh||sh.getLastRow()<2) return [];
        return parseSheet(ss,SH.PAYMENTS);
      })
    };
  } catch(e){ return { feeStructure:[], payments:[] }; }
}
