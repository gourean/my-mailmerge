<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', sans-serif; padding: 12px; font-size: 13px; color: #3c4043; margin-bottom: 60px; }
    .help-box { background: #fef7e0; border: 1px solid #feefc3; padding: 10px; border-radius: 4px; margin-bottom: 15px; }
    label { font-weight: bold; display: block; margin-top: 10px; }
    input, select, textarea { width: 100%; padding: 8px; margin-top: 4px; border: 1px solid #dadce0; border-radius: 4px; box-sizing: border-box; }
    textarea { white-space: pre-wrap; }
    .row { display: flex; gap: 8px; }
    .col { flex: 1; }
    .footer-status { position: fixed; bottom: 0; left: 0; width: 100%; background: #3c4043; color: white; padding: 10px; text-align: center; display: none; }
    .btn-blue { background: #1a73e8; color: white; border: none; padding: 10px; width: 100%; border-radius: 4px; margin-top: 10px; cursor: pointer; font-weight: bold; }
    .btn-outline { background: #fff; color: #1a73e8; border: 1px solid #dadce0; padding: 8px; width: 100%; border-radius: 4px; margin-top: 5px; cursor: pointer; }
    .instruction { font-size: 11px; color: #70757a; margin-top: 4px; font-style: italic; }
  </style>
</head>
<body>

  <div class="help-box">
    <b>💡 Personalize:</b> Use <code>{{Column Name}}</code>. Press <b>Enter</b> for new paragraphs.
  </div>

  <label>Target Email Column</label>
  <select id="emailCol"><option value="">-- Select --</option></select>

  <label>File Name Column</label>
  <select id="fileCol"><option value="">-- Select --</option></select>
  <div class="instruction">Include extension (e.g. file.pdf)</div>

  <label>Drive Folder ID</label>
  <input type="text" id="folderId" placeholder="Paste ID from URL">

  <div class="row">
    <div class="col"><label>CC</label><input type="text" id="cc"></div>
    <div class="col"><label>BCC</label><input type="text" id="bcc"></div>
  </div>

  <label>Reply-To Email</label>
  <input type="text" id="replyTo" placeholder="support@yourcompany.com">

  <label>Subject Line</label>
  <input type="text" id="subject">

  <label>Message Body</label>
  <textarea id="message" rows="8" placeholder="Type your message here..."></textarea>

  <button class="btn-blue" onclick="run(false)">🚀 RUN MY MAIL MERGE</button>
  
  <div class="row">
    <button class="btn-outline" onclick="run(true)">🧪 SEND TEST EMAIL</button>
    <button class="btn-outline" onclick="save()">💾 SAVE SETTINGS</button>
  </div>

  <div id="statusFooter" class="footer-status">PROGRESS: <span id="prog">0/0</span></div>

  <script>
    window.onload = () => {
      google.script.run.withSuccessHandler(headers => {
        ['emailCol', 'fileCol'].forEach(id => {
          const sel = document.getElementById(id);
          headers.forEach(h => sel.add(new Option(h, h)));
        });
        
        google.script.run.withSuccessHandler(saved => {
          if (saved) {
            document.getElementById('emailCol').value = saved.emailCol || "";
            document.getElementById('fileCol').value = saved.fileCol || "";
            document.getElementById('folderId').value = saved.folderId || "";
            document.getElementById('cc').value = saved.cc || "";
            document.getElementById('bcc').value = saved.bcc || "";
            document.getElementById('replyTo').value = saved.replyTo || "";
            document.getElementById('subject').value = saved.subjectLine || "";
            document.getElementById('message').value = saved.messageBody || "";
          }
        }).loadSettings();
      }).getSheetHeaders();
    };

    function save() {
      const config = {
        emailCol: document.getElementById('emailCol').value,
        fileCol: document.getElementById('fileCol').value,
        replyTo: document.getElementById('replyTo').value,
        cc: document.getElementById('cc').value,
        bcc: document.getElementById('bcc').value,
        subjectLine: document.getElementById('subject').value,
        messageBody: document.getElementById('message').value,
        folderId: document.getElementById('folderId').value,
        mode: 'attach'
      };
      google.script.run.withSuccessHandler(msg => alert(msg)).saveSettings(config);
    }

    function run(isTest) {
      document.getElementById('statusFooter').style.display = 'block';
      const config = {
        emailCol: document.getElementById('emailCol').value,
        fileCol: document.getElementById('fileCol').value,
        replyTo: document.getElementById('replyTo').value,
        cc: document.getElementById('cc').value,
        bcc: document.getElementById('bcc').value,
        subjectLine: document.getElementById('subject').value,
        messageBody: document.getElementById('message').value,
        folderId: document.getElementById('folderId').value,
        mode: 'attach',
        isTest: isTest
      };

      const timer = setInterval(() => {
        google.script.run.withSuccessHandler(p => document.getElementById('prog').innerText = p).getProgress();
      }, 1000);

      google.script.run.withSuccessHandler(msg => { 
        clearInterval(timer); 
        alert(msg); 
        document.getElementById('statusFooter').style.display = 'none';
      }).runMerge(config);
    }
  </script>
</body>
</html>
