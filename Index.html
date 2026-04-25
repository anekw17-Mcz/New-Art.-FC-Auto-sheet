<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <script src="https://unpkg.com/html5-qrcode"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;700;800&display=swap');
        body { background: #f8fafc; color: #334155; font-family: 'Prompt', sans-serif; padding: 10px; min-height: 100vh; }
        .main-card { background: #ffffff; border-radius: 35px; max-width: 500px; margin: 30px auto; overflow: hidden; box-shadow: 0 30px 60px rgba(0, 0, 0, 0.08); animation: slideUp 0.8s ease-out; }
        .glow-wrapper { position: relative; padding: 4px; border-radius: 35px 35px 0 0; overflow: hidden; background: #fff; }
        .glow-wrapper::before { content: ''; position: absolute; top: -50%; left: -50%; width: 200%; height: 200%; background: conic-gradient(transparent, #3b82f6, #8b5cf6, transparent 30%); animation: rotateGlow 4s linear infinite; }
        .header-content { position: relative; background: #fff; border-radius: 32px 32px 0 0; padding: 40px 20px; text-align: center; z-index: 1; }
        .header-title { font-weight: 800; font-size: 1.8rem; letter-spacing: 1px; margin: 0; background: linear-gradient(90deg, #1e40af, #3b82f6); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
        @keyframes rotateGlow { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        .section-box { padding: 20px 25px; border-bottom: 1px solid #f1f5f9; }
        .label-title { font-size: 0.75rem; font-weight: 700; text-transform: uppercase; color: #64748b; margin-bottom: 12px; display: block; }
        .info-display { background: #f8fafc; padding: 15px; border-radius: 18px; color: #1e293b; font-weight: 600; border: 1px solid #e2e8f0; min-height: 50px; }
        .btn-group-custom { display: flex; gap: 10px; }
        .option-btn { flex: 1; padding: 15px; border-radius: 18px; border: 2px solid #e2e8f0; background: #fff; color: #64748b; font-weight: 700; transition: all 0.3s; }
        .option-btn.active { border-color: #3b82f6; background: #eff6ff; color: #1d4ed8; box-shadow: 0 0 15px rgba(59, 130, 246, 0.3); transform: translateY(-2px); }
        .submit-btn { background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%); color: white; width: 100%; padding: 22px; border: none; border-radius: 25px; font-weight: 800; font-size: 1.2rem; box-shadow: 0 15px 35px rgba(37, 99, 235, 0.3); transition: 0.4s; }
        #loadingOverlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(255,255,255,0.9); backdrop-filter: blur(10px); z-index: 9999; text-align: center; padding-top: 60%; }
        @keyframes slideUp { from { opacity: 0; transform: translateY(40px); } to { opacity: 1; transform: translateY(0); } }
    </style>
</head>
<body>
    <div id="loadingOverlay">
        <div class="spinner-border text-primary" style="width: 3rem; height: 3rem;"></div>
        <div class="mt-3 fw-bold">ระบบกำลังบันทึกข้อมูล...</div>
    </div>

    <div class="main-card">
        <div class="glow-wrapper"><div class="header-content"><h1 class="header-title">NEW ART. FC REPORT</h1></div></div>
        <form id="inspectionForm">
            <div class="section-box">
                <label class="label-title">Article / Barcode</label>
                <div class="input-group">
                    <input type="number" id="article" class="form-control" onchange="fetchData()" placeholder="Scan Here" style="border-radius: 18px 0 0 18px; padding: 14px;">
                    <button class="btn btn-dark px-4" type="button" onclick="startScan()" style="border-radius: 0 18px 18px 0;"><i class="fas fa-barcode"></i></button>
                </div>
                <div id="reader" style="display:none; margin-top:15px; border-radius:25px; overflow:hidden;"></div>
                <div id="reader-error" style="color: red; font-size: 0.8rem; margin-top: 5px; display: none;"></div>
            </div>

            <div class="section-box bg-light bg-opacity-50">
                <div class="row g-3">
                    <div class="col-12"><label class="label-title">Vendor Name</label><div id="v_vendor" class="info-display"></div></div>
                    <div class="col-12"><label class="label-title">Article Description</label><div id="v_desc" class="info-display"></div></div>
                    <div class="col-12"><label class="label-title">PO Number</label><div id="v_po" class="info-display"></div></div>
                    <div class="col-6"><label class="label-title">PO Qty</label><div id="v_qty" class="info-display"></div></div>
                    <div class="col-6"><label class="label-title">Sampling</label><div id="v_samp" class="info-display"></div></div>
                    <div class="col-6"><label class="label-title">Date</label><div id="v_date" class="info-display"></div></div>
                    <div class="col-6"><label class="label-title">วันผลิต (MFD)</label><input type="text" id="mfgDate" class="form-control" placeholder="ว/ด/ป" style="border-radius: 18px; padding: 14px;"></div>
                </div>
            </div>

            <div class="section-box">
                <label class="label-title" style="color:#0ea5e9;">ตัวอย่างสินค้า</label>
                <div class="btn-group-custom">
                    <button type="button" class="option-btn active" onclick="setOption('sampleStatus', 'มีตัวอย่าง', this)">มีตัวอย่าง</button>
                    <button type="button" class="option-btn" onclick="setOption('sampleStatus', 'ไม่มีตัวอย่าง', this)">ไม่มีตัวอย่าง</button>
                </div>
                <input type="hidden" id="sampleStatus" value="มีตัวอย่าง">
            </div>

            <div class="section-box">
                <label class="label-title" style="color:#f59e0b;">ผลการตรวจสอบ</label>
                <div class="btn-group-custom">
                    <button type="button" class="option-btn active" onclick="setOption('inspectionResult', 'ไม่พบปัญหา', this)">ไม่พบปัญหา</button>
                    <button type="button" class="option-btn" onclick="setOption('inspectionResult', 'พบปัญหา', this)">พบปัญหา</button>
                </div>
                <input type="hidden" id="inspectionResult" value="ไม่พบปัญหา">
            </div>

            <div class="section-box border-0">
                <label class="label-title">1. รูปหน้ากล่อง/หลังกระเบื้อง/มอก. และสคบ.</label>
                <input type="file" id="f1" class="form-control mb-3" multiple accept="image/*" style="border-radius:18px;">
                <label class="label-title">2. รูปภาพโดยรวมมุมไกล</label>
                <input type="file" id="f2" class="form-control mb-3" multiple accept="image/*" style="border-radius:18px;">
                <label class="label-title">3. รูปภาพโดยรวมมุมใกล้</label>
                <input type="file" id="f3" class="form-control mb-4" multiple accept="image/*" style="border-radius:18px;">

                <button type="button" class="submit-btn" onclick="submitForm()" id="submitBtn">บันทึกข้อมูลและส่งรายงาน</button>
            </div>
        </form>
    </div>

    <input type="hidden" id="h_date"><input type="hidden" id="h_po"><input type="hidden" id="h_vendor"><input type="hidden" id="h_desc"><input type="hidden" id="h_qty"><input type="hidden" id="h_samp">

    <script>
        function setOption(targetId, value, btnElement) {
            document.getElementById(targetId).value = value;
            btnElement.parentElement.querySelectorAll('.option-btn').forEach(btn => btn.classList.remove('active'));
            btnElement.classList.add('active');
        }

        async function compressImage(file) {
            return new Promise((resolve) => {
                const reader = new FileReader();
                reader.readAsDataURL(file);
                reader.onload = (e) => {
                    const img = new Image();
                    img.src = e.target.result;
                    img.onload = () => {
                        const canvas = document.createElement('canvas');
                        let width = img.width, height = img.height;
                        const maxPixels = 900000;
                        if (width * height > maxPixels) {
                            const ratio = Math.sqrt(maxPixels / (width * height));
                            width = Math.floor(width * ratio);
                            height = Math.floor(height * ratio);
                        }
                        if (width > 1000) { height = Math.floor(height * (1000 / width)); width = 1000; }
                        canvas.width = width; canvas.height = height;
                        canvas.getContext('2d').drawImage(img, 0, 0, width, height);
                        resolve(canvas.toDataURL('image/jpeg', 0.5)); // บีบอัดรูปป้องกัน Error
                    };
                };
            });
        }

        async function getGroupImages(id) {
            const files = document.getElementById(id).files;
            if (files.length === 0) return null;
            const promises = Array.from(files).map(async (file) => {
                const base64 = await compressImage(file);
                return { data: base64.split(',')[1], type: 'image/jpeg', name: file.name };
            });
            return await Promise.all(promises);
        }

        function fetchData() {
            var inputVal = document.getElementById('article').value;
            if (!inputVal) return;
            google.script.run.withSuccessHandler(data => {
                if (data) {
                    // เปลี่ยนค่า Barcode เป็นเลข Article
                    document.getElementById('article').value = data.article;
                    
                    // แสดงข้อมูลให้ตรงช่อง
                    document.getElementById('v_vendor').innerText = data.vendor;
                    document.getElementById('v_desc').innerText = data.description;
                    document.getElementById('v_po').innerText = data.poNumber;
                    document.getElementById('v_qty').innerText = data.poQty;
                    document.getElementById('v_samp').innerText = data.sampling;
                    document.getElementById('v_date').innerText = data.date;

                    document.getElementById('h_vendor').value = data.vendor;
                    document.getElementById('h_po').value = data.poNumber;
                    document.getElementById('h_date').value = data.date;
                    document.getElementById('h_desc').value = data.description;
                    document.getElementById('h_qty').value = data.poQty;
                    document.getElementById('h_samp').value = data.sampling;
                } else {
                    alert("ไม่พบข้อมูลสินค้าจากบาร์โค้ดนี้!");
                }
            }).getProductData(inputVal);
        }

        async function submitForm() {
            document.getElementById('loadingOverlay').style.display = 'block';
            const btn = document.getElementById('submitBtn'); btn.disabled = true;
            try {
                const payload = {
                    article: document.getElementById('article').value,
                    date: document.getElementById('h_date').value,
                    poNumber: document.getElementById('h_po').value,
                    vendor: document.getElementById('h_vendor').value,
                    description: document.getElementById('h_desc').value,
                    poQty: document.getElementById('h_qty').value,
                    sampling: document.getElementById('h_samp').value,
                    mfgDate: document.getElementById('mfgDate').value,
                    inspectionResult: document.getElementById('inspectionResult').value,
                    sampleStatus: document.getElementById('sampleStatus').value,
                    group1: await getGroupImages('f1'),
                    group2: await getGroupImages('f2'),
                    group3: await getGroupImages('f3')
                };
                google.script.run.withSuccessHandler(function(msg) {
                    alert(msg);
                    // ล้างข้อมูลเพื่อเริ่ม Article ใหม่ (ป้องกันหน้าขาว)
                    document.getElementById('inspectionForm').reset();
                    ['v_vendor','v_desc','v_po','v_qty','v_samp','v_date'].forEach(id => document.getElementById(id).innerText = "");
                    document.querySelectorAll('.option-btn').forEach(b => b.classList.remove('active'));
                    document.querySelectorAll('.btn-group-custom .option-btn:first-child').forEach(b => b.classList.add('active'));
                    document.getElementById('sampleStatus').value = "มีตัวอย่าง";
                    document.getElementById('inspectionResult').value = "ไม่พบปัญหา";
                    document.getElementById('loadingOverlay').style.display = 'none';
                    btn.disabled = false;
                    window.scrollTo(0, 0);
                }).processForm(payload);
            } catch (e) { alert("Error: " + e); document.getElementById('loadingOverlay').style.display = 'none'; btn.disabled = false; }
        }

        function startScan() {
            const readerElement = document.getElementById('reader');
            const errorElement = document.getElementById('reader-error');
            const html5QrCode = new Html5Qrcode("reader");
            readerElement.style.display = 'block';
            errorElement.style.display = 'none';

            const qrConfig = { fps: 10, qrbox: { width: 250, height: 180 }, aspectRatio: 1.0 };

            html5QrCode.start({ facingMode: "environment" }, qrConfig, 
                (decodedText) => {
                    document.getElementById('article').value = decodedText;
                    fetchData();
                    html5QrCode.stop();
                    readerElement.style.display = 'none';
                },
                (errorMessage) => {}
            ).catch(err => {
                readerElement.style.display = 'none';
                errorElement.innerText = "ไม่สามารถเข้าถึงกล้องได้: โปรดเปิดลิงก์ใน Chrome/Safari";
                errorElement.style.display = 'block';
                alert("เกิดข้อผิดพลาด: " + err + "\n\nกรุณาเปิดด้วย Chrome หรือ Safari (ไม่ใช่เบราว์เซอร์ใน Line)");
            });
        }
    </script>
</body>
</html>
