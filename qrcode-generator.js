document.addEventListener('DOMContentLoaded', function() {
    const contentType = document.getElementById('content-type');
    const content = document.getElementById('content');
    const size = document.getElementById('size');
    const foreground = document.getElementById('foreground');
    const background = document.getElementById('background');
    const generateBtn = document.getElementById('generate-btn');
    const downloadBtn = document.getElementById('download-btn');
    const scanBtn = document.getElementById('scan-btn');
    const qrcode = document.getElementById('qrcode');
    const scanPreview = document.getElementById('scan-preview');
    
    let qrCanvas = null;
    
    // 生成二维码
    function generateQRCode() {
        const type = contentType.value;
        let contentValue = content.value.trim();
        
        if (!contentValue) {
            alert('请输入内容');
            return;
        }
        
        // 根据类型处理内容
        switch(type) {
            case 'url':
                if (!contentValue.startsWith('http://') && !contentValue.startsWith('https://')) {
                    contentValue = 'https://' + contentValue;
                }
                break;
            case 'phone':
                contentValue = 'tel:' + contentValue;
                break;
            case 'email':
                contentValue = 'mailto:' + contentValue;
                break;
        }
        
        // 清空之前的二维码
        qrcode.innerHTML = '';
        
        // 创建Canvas元素
        const canvas = document.createElement('canvas');
        qrCanvas = canvas;
        qrcode.appendChild(canvas);
        
        // 生成二维码
        QRCode.toCanvas(canvas, contentValue, {
            width: parseInt(size.value),
            margin: 1,
            color: {
                dark: foreground.value,
                light: background.value
            }
        }, function(error) {
            if (error) {
                console.error(error);
                alert('生成二维码失败');
            }
        });
    }
    
    // 下载二维码为PNG
    function downloadQRCode() {
        if (!qrCanvas) {
            alert('请先生成二维码');
            return;
        }
        
        // 将Canvas转换为PNG图片
        const dataURL = qrCanvas.toDataURL('image/png');
        
        // 创建下载链接
        const link = document.createElement('a');
        link.download = 'qrcode.png';
        link.href = dataURL;
        link.click();
    }
    
    // 扫描预览二维码内容
    function scanQRCode() {
        if (!qrCanvas) {
            alert('请先生成二维码');
            return;
        }
        
        // 获取Canvas的图像数据
        const ctx = qrCanvas.getContext('2d');
        const imageData = ctx.getImageData(0, 0, qrCanvas.width, qrCanvas.height);
        
        // 扫描二维码
        const code = jsQR(imageData.data, imageData.width, imageData.height);
        
        if (code) {
            scanPreview.textContent = '扫描结果: ' + code.data;
        } else {
            scanPreview.textContent = '扫描失败，无法识别二维码';
        }
    }
    
    // 事件监听
    generateBtn.addEventListener('click', generateQRCode);
    downloadBtn.addEventListener('click', downloadQRCode);
    scanBtn.addEventListener('click', scanQRCode);
    
    // 内容变化时自动生成
    content.addEventListener('input', generateQRCode);
    contentType.addEventListener('change', generateQRCode);
    size.addEventListener('change', generateQRCode);
    foreground.addEventListener('change', generateQRCode);
    background.addEventListener('change', generateQRCode);
});