// 等待文档加载完毕
document.addEventListener('DOMContentLoaded', () => {

    // 获取页面上的元素
    const promptTemplateInput = document.getElementById('prompt-template');
    const excelFileInput = document.getElementById('excel-file-input');
    const generateBtn = document.getElementById('generate-btn');
    const statusMessage = document.getElementById('status-message');
    const downloadLink = document.getElementById('download-link');
    const outputPreviewPanel = document.querySelector('.output-preview-panel');
    const outputPreview = document.getElementById('output-preview');

    // 为“生成”按钮添加点击事件监听
    generateBtn.addEventListener('click', () => {
        // 获取用户输入
        const promptTemplate = promptTemplateInput.value.trim();
        const file = excelFileInput.files[0];

        // 1. 输入校验
        if (!promptTemplate) {
            updateStatus('错误：Prompt 模板不能为空！', 'error');
            return;
        }
        if (!file) {
            updateStatus('错误：请上传一个 Excel 文件！', 'error');
            return;
        }

        // 重置状态
        updateStatus('正在处理文件...', 'info');
        hideOutput();

        // 2. 读取和解析 Excel 文件
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // 将工作表转换为 JSON 对象数组，第一行作为键
                const jsonData = XLSX.utils.sheet_to_json(worksheet);

                if (jsonData.length === 0) {
                    updateStatus('错误：Excel 文件是空的或格式不正确。', 'error');
                    return;
                }
                
                // 3. 生成 JSONL 内容
                const jsonlContent = generateJsonl(promptTemplate, jsonData);

                if (jsonlContent) {
                    // 4. 创建并提供下载链接
                    createDownloadLink(jsonlContent);
                    // 5. 显示预览
                    showPreview(jsonlContent);
                    updateStatus(`成功生成 ${jsonData.length} 条记录！`, 'success');
                }
            } catch (error) {
                console.error(error);
                updateStatus(`处理失败：${error.message}`, 'error');
            }
        };

        reader.onerror = () => {
            updateStatus('读取文件失败！', 'error');
        };

        reader.readAsArrayBuffer(file);
    });

    /**
     * 根据模板和数据生成 JSONL 字符串
     * @param {string} template - Prompt 模板
     * @param {Array<Object>} data - 从 Excel 解析出的数据数组
     * @returns {string|null} - 生成的 JSONL 字符串或 null
     */
    function generateJsonl(template, data) {
        const lines = [];
        const headers = Object.keys(data[0]); // 获取表头

        // 检查模板中的变量是否都存在于Excel表头中
        const templateVars = template.match(/\{(.+?)\}/g) || [];
        const missingVars = templateVars
            .map(v => v.slice(1, -1)) // 去掉花括号
            .filter(v => !headers.includes(v));

        if (missingVars.length > 0) {
            updateStatus(`错误：模板中的变量 {${missingVars.join(', ')}} 在 Excel 表头中找不到！`, 'error');
            return null;
        }

        data.forEach(row => {
            let filledPrompt = template;
            // 遍历每一列，替换模板中的变量
            for (const key in row) {
                const regex = new RegExp(`\\{${key}\\}`, 'g');
                filledPrompt = filledPrompt.replace(regex, row[key]);
            }
            
            // 构建符合通用格式的 JSON 对象
            const jsonObject = {
                // 这里可以自定义你想要的 JSON 结构
                // 常见的是 messages 格式，用于聊天模型
                "prompt": filledPrompt
            };

            lines.push(JSON.stringify(jsonObject));
        });

        return lines.join('\n');
    }
    
    /**
     * 创建下载链接
     * @param {string} content - 文件内容
     */
    function createDownloadLink(content) {
        const blob = new Blob([content], { type: 'application/jsonl' });
        const url = URL.createObjectURL(blob);
        downloadLink.href = url;
        downloadLink.download = 'output.jsonl';
        downloadLink.classList.remove('hidden');
    }

    /**
     * 更新状态消息
     * @param {string} message - 要显示的消息
     * @param {string} type - 消息类型 ('info', 'success', 'error')
     */
    function updateStatus(message, type = 'info') {
        statusMessage.textContent = message;
        statusMessage.style.color = type === 'error' ? '#d9534f' : (type === 'success' ? '#5cb85c' : '#333');
    }
    
    /**
     * 显示预览
     * @param {string} content - JSONL 全部内容
     */
    function showPreview(content) {
        const previewLines = content.split('\n').slice(0, 5).join('\n');
        outputPreview.textContent = previewLines;
        outputPreviewPanel.classList.remove('hidden');
    }

    /**
     * 隐藏输出区域
     */
    function hideOutput() {
        downloadLink.classList.add('hidden');
        outputPreviewPanel.classList.add('hidden');
        outputPreview.textContent = '';
    }
});
