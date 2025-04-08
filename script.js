// 動態生成10個問題
document.addEventListener('DOMContentLoaded', () => {
    const questionsDiv = document.getElementById('questions');
    for (let i = 1; i <= 10; i++) {
        const div = document.createElement('div');
        div.className = 'question';
        div.innerHTML = `
            <label>問題 ${i}: </label>
            <input type="text" id="q${i}" placeholder="請輸入答案">
            <input type="file" id="photo${i}" accept="image/*">
        `;
        questionsDiv.appendChild(div);
    }
});

// 同時匯出 Excel 和 PDF
function exportBoth() {
    try {
        // 檢查必要庫是否載入
        if (typeof XLSX === 'undefined') throw new Error('SheetJS 未載入，請檢查網路或使用本地檔案');
        if (typeof window.jspdf === 'undefined') throw new Error('jsPDF 未載入，請檢查網路或使用本地檔案');
        if (typeof NotoSansTCRegular === 'undefined') {
            console.warn('Noto Sans TC 字型未載入，將使用預設字型（可能導致中文亂碼）');
        }

        // 收集資料
        const data = [];
        const promises = [];
        for (let i = 1; i <= 10; i++) {
            const answer = document.getElementById(`q${i}`).value || "未填寫";
            const photoInput = document.getElementById(`photo${i}`);
            const photoName = photoInput.files[0] ? photoInput.files[0].name : "無相片";
            data.push([`問題 ${i}`, answer, photoName]);

            if (photoInput.files[0]) {
                promises.push(new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve({ index: i, imgData: reader.result });
                    reader.onerror = () => reject(new Error(`圖片 ${i} 讀取失敗`));
                    reader.readAsDataURL(photoInput.files[0]);
                }));
            }
        }

        // 生成 Excel
        const ws = XLSX.utils.aoa_to_sheet([['問題', '答案', '相片路徑'], ...data]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '稽核結果');
        XLSX.writeFile(wb, 'audit_result.xlsx');

        // 生成 PDF
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        if (typeof NotoSansTCRegular !== 'undefined') {
            doc.addFileToVFS('NotoSansTC-Regular.ttf', NotoSansTCRegular);
            doc.addFont('NotoSansTC-Regular.ttf', 'NotoSansTC', 'normal');
            doc.setFont('NotoSansTC');
        } else {
            console.warn('使用預設字型，可能無法正確顯示中文');
        }
        let y = 10;

        Promise.all(promises).then((images) => {
            for (let i = 1; i <= 10; i++) {
                const answer = document.getElementById(`q${i}`).value || "未填寫";
                doc.text(`問題 ${i}: ${answer}`, 10, y);
                y += 10;

                const img = images.find(img => img.index === i);
                if (img) {
                    try {
                        doc.addImage(img.imgData, 'JPEG', 10, y, 50, 50);
                        y += 60;
                    } catch (imgError) {
                        console.warn(`圖片 ${i} 添加失敗:`, imgError);
                        y += 10;
                    }
                } else {
                    y += 10;
                }
            }
            doc.save('audit_result.pdf');

            // 觸發郵件客戶端
            const subject = encodeURIComponent('稽核結果');
            const body = encodeURIComponent('請查看附件中的稽核結果（audit_result.xlsx 和 audit_result.pdf）');
            const mailtoLink = `mailto:your-email@example.com?subject=${subject}&body=${body}`;
            window.location.href = mailtoLink;
        }).catch((error) => {
            console.error('PDF 生成失敗:', error);
            alert('PDF 生成失敗，但 Excel 已匯出');
        });

    } catch (error) {
        console.error('匯出失敗:', error);
        alert('匯出時發生錯誤，請檢查控制台訊息');
    }
}