<!DOCTYPE html>
<html lang="zh-CN">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>年报填报系统</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Font Awesome 图标 -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- 自定义样式 -->
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        primary: {
                            50: '#eff6ff',
                            500: '#3b82f6',
                            600: '#2563eb',
                            700: '#1d4ed8',
                        }
                    }
                }
            }
        }
    </script>
</head>

<body class="bg-gradient-to-br from-blue-50 to-indigo-100 min-h-screen">
    <!-- 导出完成/失败时由 iframe 内 postMessage 触发，避免隐藏 iframe 内提示被忽略 -->
    <div id="exportToast" class="hidden fixed top-4 left-1/2 -translate-x-1/2 z-[100] w-[min(92vw,28rem)] rounded-xl shadow-2xl border overflow-hidden" role="status" aria-live="polite">
        <div id="exportToastInner" class="px-5 py-4">
            <p id="exportToastTitle" class="font-bold text-base"></p>
            <p id="exportToastMsg" class="text-sm mt-2 leading-relaxed opacity-95 whitespace-pre-line"></p>
        </div>
        <button type="button" id="exportToastClose" class="absolute top-2 right-2 w-8 h-8 rounded-lg hover:bg-black/10 text-lg leading-none" aria-label="关闭">×</button>
    </div>
    <!-- 主要内容区域：减小左右留白，卡片略加宽以更贴近窗口两侧 -->
    <div class="w-full max-w-7xl mx-auto px-2 sm:px-3 md:px-4 py-8">
        <!-- 页面标题 -->
        <div class="text-center mb-8">
            <h4 class="text-3xl font-bold text-gray-800 mb-2">
                <i class="fas fa-chart-line text-primary-600 mr-3"></i>
                年报填报系统
            </h4>
        </div>

        <!-- 表单卡片（原 max-w-2xl 较窄，大屏两侧空白多） -->
        <div class="max-w-5xl mx-auto">
            <div class="bg-white rounded-xl shadow-xl p-8 border border-gray-100">
                <!-- 双上传区域：默认最窄视口单列，sm 及以上双列（避免 lg 前长时间单列） -->
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-8">
                    <!-- 教师导入区域 -->
                    <div class="bg-gradient-to-br from-blue-50 to-blue-100 rounded-xl p-6 border border-blue-200">
                        <!-- 小标题与图标同一行 -->
                        <div class="flex items-center justify-center gap-3 mb-6">
                            <div class="w-12 h-12 shrink-0 bg-blue-500 rounded-full flex items-center justify-center" aria-hidden="true">
                                <i class="fas fa-chalkboard-teacher text-white text-xl"></i>
                            </div>
                            <h3 class="text-xl font-bold text-blue-800 m-0">教师信息导入</h3>
                        </div>

                        <!-- 模板位于 www 目录，直链下载（文件名含中文需 URL 编码） -->
                        <p class="text-center text-sm mb-4">
                            <a href="<?php echo rawurlencode('教职工信息表.xlsx'); ?>"
                               class="inline-flex items-center justify-center text-blue-700 hover:text-blue-900 font-medium underline decoration-blue-400 underline-offset-2"
                               download>
                                <i class="fas fa-file-download mr-2 text-blue-600" aria-hidden="true"></i>
                                下载模板：教职工信息表.xlsx
                            </a>
                        </p>

                        <form action="?action=import_teachers" method="post" enctype="multipart/form-data" class="space-y-4" id="teacherUploadForm">

                            <!-- 文件选择区域 -->
                            <div class="space-y-3">
                               

                                <!-- 拖拽上传区域 -->
                                <div class="relative">
                                    <div class="border-2 border-dashed border-blue-300 rounded-lg p-6 text-center hover:border-blue-400 transition-colors duration-200 bg-white" id="teacherDropZone">
                                        <div class="space-y-3">
                                            <div class="mx-auto w-12 h-12 bg-blue-50 rounded-full flex items-center justify-center">
                                                <i class="fas fa-cloud-upload-alt text-blue-500 text-lg"></i>
                                            </div>
                                            <div>
                                                <p class="text-sm font-medium text-gray-700">选择或拖拽文件</p>
                                                <p class="text-xs text-gray-500 mt-1">支持 .xls 和 .xlsx 格式</p>
                                            </div>
                                            <!-- accept 仅用扩展名：Windows/CEF 打开文件对话框时更易默认筛选为 xls、xlsx（MIME 混写有时落到「所有文件」） -->
                                            <input type="file"
                                                name="file"
                                                id="teacherFile"
                                                accept=".xlsx,.xls"
                                                class="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                                                required>
                                        </div>
                                    </div>

                                    <!-- 文件信息显示 -->
                                    <div id="teacherFileInfo" class="hidden mt-3 p-3 bg-green-50 border border-green-200 rounded-lg">
                                        <div class="flex items-center">
                                            <i class="fas fa-check-circle text-green-500 mr-2"></i>
                                            <span class="text-green-700 font-medium text-sm" id="teacherFileName"></span>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <!-- 提交按钮 -->
                            <button type="submit"
                                name="submit"
                                value="submit"
                                class="w-full bg-gradient-to-r from-blue-500 to-blue-600 hover:from-blue-600 hover:to-blue-700 text-white font-semibold py-3 px-4 rounded-lg shadow-lg transform hover:scale-105 transition-all duration-200 focus:outline-none focus:ring-4 focus:ring-blue-200">
                                <i class="fas fa-download mr-2"></i>
                                生成教职工报表
                            </button>

                            <!-- 生成报表说明 -->
                            <div class="mt-4 p-3 bg-blue-50 border border-blue-200 rounded-lg">
                                <p class="text-sm text-blue-800 font-medium mb-2">
                                    <i class="fas fa-info-circle mr-1"></i>
                                    可生成以下报表：
                                </p>
                                <ul class="text-xs text-blue-700 space-y-1 ml-5">
                                    <li>• 教基4149 中小学教职工</li>
                                    <li>• 教基4153 专任教师分年龄情况</li>
                                    <li>• 教基4155 专任教师分课程、分学历情况</li>
                                    <li>• 教基4067 教职工其他情况</li>
                                </ul>
                            </div>

                        </form>
                    </div>

                    <!-- 学生导入区域 -->
                    <div class="bg-gradient-to-br from-green-50 to-green-100 rounded-xl p-6 border border-green-200">
                        <!-- 小标题与图标同一行 -->
                        <div class="flex items-center justify-center gap-3 mb-6">
                            <div class="w-12 h-12 shrink-0 bg-green-500 rounded-full flex items-center justify-center" aria-hidden="true">
                                <i class="fas fa-user-graduate text-white text-xl"></i>
                            </div>
                            <h3 class="text-xl font-bold text-green-800 m-0">学生信息导入</h3>
                        </div>

                        <p class="text-center text-sm mb-4">
                            <a href="<?php echo rawurlencode('学生信息表.xlsx'); ?>"
                               class="inline-flex items-center justify-center text-green-700 hover:text-green-900 font-medium underline decoration-green-400 underline-offset-2"
                               download>
                                <i class="fas fa-file-download mr-2 text-green-600" aria-hidden="true"></i>
                                下载模板：学生信息表.xlsx
                            </a>
                        </p>

                        <form action="?action=import_students" method="post" enctype="multipart/form-data" class="space-y-4" id="studentUploadForm">

                            <!-- 文件选择区域 -->
                            <div class="space-y-3">
                              

                                <!-- 拖拽上传区域 -->
                                <div class="relative">
                                    <div class="border-2 border-dashed border-green-300 rounded-lg p-6 text-center hover:border-green-400 transition-colors duration-200 bg-white" id="studentDropZone">
                                        <div class="space-y-3">
                                            <div class="mx-auto w-12 h-12 bg-green-50 rounded-full flex items-center justify-center">
                                                <i class="fas fa-cloud-upload-alt text-green-500 text-lg"></i>
                                            </div>
                                            <div>
                                                <p class="text-sm font-medium text-gray-700">选择或拖拽文件</p>
                                                <p class="text-xs text-gray-500 mt-1">支持 .xls 和 .xlsx 格式</p>
                                            </div>
                                            <!-- 与教师上传一致：对话框默认仅突出 Excel 扩展名 -->
                                            <input type="file"
                                                name="file"
                                                id="studentFile"
                                                accept=".xlsx,.xls"
                                                class="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                                                required>
                                        </div>
                                    </div>

                                    <!-- 文件信息显示 -->
                                    <div id="studentFileInfo" class="hidden mt-3 p-3 bg-green-50 border border-green-200 rounded-lg">
                                        <div class="flex items-center">
                                            <i class="fas fa-check-circle text-green-500 mr-2"></i>
                                            <span class="text-green-700 font-medium text-sm" id="studentFileName"></span>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <!-- 提交按钮 -->
                            <button type="submit"
                                name="submit"
                                value="submit"
                                class="w-full bg-gradient-to-r from-green-500 to-green-600 hover:from-green-600 hover:to-green-700 text-white font-semibold py-3 px-4 rounded-lg shadow-lg transform hover:scale-105 transition-all duration-200 focus:outline-none focus:ring-4 focus:ring-green-200">
                                <i class="fas fa-download mr-2"></i>
                                生成学生报表
                            </button>

                            <!-- 生成报表说明 -->
                            <div class="mt-4 p-3 bg-green-50 border border-green-200 rounded-lg">
                                <p class="text-sm text-green-800 font-medium mb-2">
                                    <i class="fas fa-info-circle mr-1"></i>
                                    可生成以下报表：
                                </p>
                                <ul class="text-xs text-green-700 space-y-1 ml-5">
                                    <li>• 教基3112 小学分类型学生数</li>
                                    <li>• 教基3113 小学分年龄学生数</li>
                                    <li>• 教基3115 初中分类型学生数</li>
                                    <li>• 教基3116 初中分年龄学生数</li>
                                    <li>• 教基3118 普通高中分类型学生数</li>
                                    <li>• 教基3119 高中阶段分年龄学生数</li>
                                </ul>
                            </div>

                        </form>
                    </div>
                </div>

                <!-- 提示信息 -->
                <div class="mt-4 p-4 bg-blue-50 border border-blue-200 rounded-lg">
                    <div class="flex items-start">
                        <i class="fas fa-info-circle text-blue-500 mr-2 mt-0.5"></i>
                        <div class="text-blue-700">
                            <p class="font-medium">说明：</p>
                            <ul class="text-sm mt-1 space-y-1">
                                <li>• 首先从上方下载模板文件，并按年报手册数据要求填写</li>
                                <li>• 序号和姓名两列非必须项</li>
                                <li>• 导入过程可能需要较长时间，请耐心等待</li>
                                
                                <li>• 导入过程中请勿关闭页面</li>
                                <li>• 导出结果：桌面版会保存到程序运行目录；在线版自动下载</li>
                                <li>• 部分项目因不方便自动统计或者数据量少无必要自动统计，请手动填写</li>
                                <li>• 对于非大陆身份证号码，可以按照如下规则编一个，以便自动提取出生日期和性别，六位任意数字+出生日期+任意两位数字+性别（男1女2）+一位任意数字</li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- JavaScript 功能 -->
    <script>
        // 教师文件拖拽功能
        const teacherDropZone = document.getElementById('teacherDropZone');
        const teacherFileInput = document.getElementById('teacherFile');
        const teacherFileInfo = document.getElementById('teacherFileInfo');
        const teacherFileName = document.getElementById('teacherFileName');

        // 学生文件拖拽功能
        const studentDropZone = document.getElementById('studentDropZone');
        const studentFileInput = document.getElementById('studentFile');
        const studentFileInfo = document.getElementById('studentFileInfo');
        const studentFileName = document.getElementById('studentFileName');

        // HTML5 required 默认提示为英文，改为中文（需在 change 后 setCustomValidity('') 以便再次校验）
        teacherFileInput.addEventListener('invalid', function () {
            if (teacherFileInput.validity.valueMissing) {
                teacherFileInput.setCustomValidity('请先选择要上传的教师 Excel 文件');
            }
        });
        studentFileInput.addEventListener('invalid', function () {
            if (studentFileInput.validity.valueMissing) {
                studentFileInput.setCustomValidity('请先选择要上传的学生 Excel 文件');
            }
        });

        // 教师拖拽事件处理
        teacherDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            teacherDropZone.classList.add('border-blue-400', 'bg-blue-50');
        });

        teacherDropZone.addEventListener('dragleave', (e) => {
            e.preventDefault();
            teacherDropZone.classList.remove('border-blue-400', 'bg-blue-50');
        });

        teacherDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            teacherDropZone.classList.remove('border-blue-400', 'bg-blue-50');

            const files = e.dataTransfer.files;
            if (files.length > 0) {
                teacherFileInput.files = files;
                teacherFileInput.setCustomValidity('');
                showTeacherFileInfo(files[0]);
            }
        });

        // 学生拖拽事件处理
        studentDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            studentDropZone.classList.add('border-green-400', 'bg-green-50');
        });

        studentDropZone.addEventListener('dragleave', (e) => {
            e.preventDefault();
            studentDropZone.classList.remove('border-green-400', 'bg-green-50');
        });

        studentDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            studentDropZone.classList.remove('border-green-400', 'bg-green-50');

            const files = e.dataTransfer.files;
            if (files.length > 0) {
                studentFileInput.files = files;
                studentFileInput.setCustomValidity('');
                showStudentFileInfo(files[0]);
            }
        });

        // 教师文件选择事件
        teacherFileInput.addEventListener('change', (e) => {
            teacherFileInput.setCustomValidity('');
            if (e.target.files.length > 0) {
                showTeacherFileInfo(e.target.files[0]);
            }
        });

        // 学生文件选择事件
        studentFileInput.addEventListener('change', (e) => {
            studentFileInput.setCustomValidity('');
            if (e.target.files.length > 0) {
                showStudentFileInfo(e.target.files[0]);
            }
        });

        // 显示教师文件信息
        function showTeacherFileInfo(file) {
            teacherFileName.textContent = `已选择文件：${file.name} (${formatFileSize(file.size)})`;
            teacherFileInfo.classList.remove('hidden');
        }

        // 显示学生文件信息
        function showStudentFileInfo(file) {
            studentFileName.textContent = `已选择文件：${file.name} (${formatFileSize(file.size)})`;
            studentFileInfo.classList.remove('hidden');
        }

        // 格式化文件大小
        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }

        // 导出结果提示（iframe 内 PHP 页通过 postMessage 通知父窗口）
        const exportToast = document.getElementById('exportToast');
        const exportToastInner = document.getElementById('exportToastInner');
        const exportToastTitle = document.getElementById('exportToastTitle');
        const exportToastMsg = document.getElementById('exportToastMsg');
        let exportToastTimer = null;

        function hideExportToast() {
            exportToast.classList.add('hidden');
            if (exportToastTimer) {
                clearTimeout(exportToastTimer);
                exportToastTimer = null;
            }
        }

        function showExportToast(ok, titleText, msgText) {
            exportToastTitle.textContent = titleText;
            exportToastMsg.textContent = msgText;
            exportToastInner.className = 'px-5 py-4 pr-10 ' + (ok ?
                'bg-emerald-600 text-white border-emerald-700' :
                'bg-red-600 text-white border-red-700');
            exportToast.classList.remove('hidden');
            if (exportToastTimer) clearTimeout(exportToastTimer);
            exportToastTimer = setTimeout(hideExportToast, ok ? 12000 : 16000);
        }
        document.getElementById('exportToastClose').addEventListener('click', hideExportToast);
        window.addEventListener('message', function(ev) {
            if (!ev.data || ev.data.type !== 'edureport-export') return;
            if (ev.origin !== window.location.origin) return;
            if (ev.data.ok) {
                const path = ev.data.path || '';
                const verified = ev.data.verified !== false;
                const name = ev.data.filename || '';
                const sub = verified ?
                    ('文件已保存：' + (name || path) + (path ? '\n完整路径：' + path : '')) :
                    ('未确认文件已写入，请检查路径：' + path);
                showExportToast(true, '导出完成', sub);
            } else {
                showExportToast(false, '导出失败', (ev.data.message || '未知错误') + (ev.data.path ? '\n目标：' + ev.data.path : ''));
            }
        });

        // 教师表单提交确认
        document.getElementById('teacherUploadForm').addEventListener('submit', function(e) {
            if (!teacherFileInput.files.length) {
                e.preventDefault();
                alert('请先选择要上传的教师文件');
                return false;
            }

            const submitBtn = this.querySelector('button[type="submit"]');
            const originalHtml = submitBtn.innerHTML;
            submitBtn.dataset.originalHtml = originalHtml;
            submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>正在生成教职工报表...';
            submitBtn.disabled = true;

            // 创建隐藏的iframe用于提交表单，避免页面导航
            const iframe = document.createElement('iframe');
            iframe.name = 'teacherDownloadFrame';
            iframe.style.display = 'none';
            document.body.appendChild(iframe);

            // 设置表单target为iframe
            this.target = 'teacherDownloadFrame';

            // 监听iframe加载完成（文件下载完成后）
            iframe.onload = function() {
                // 文件下载完成后恢复按钮状态
                setTimeout(function() {
                    submitBtn.innerHTML = originalHtml;
                    submitBtn.disabled = false;
                    // 移除iframe
                    document.body.removeChild(iframe);
                }, 1000);
            };

            // 备用方案：如果iframe的onload未触发（某些浏览器对文件下载的处理不同），使用定时器恢复
            setTimeout(function() {
                if (submitBtn.disabled) {
                    submitBtn.innerHTML = originalHtml;
                    submitBtn.disabled = false;
                    if (iframe.parentNode) {
                        document.body.removeChild(iframe);
                    }
                }
            }, 15000); // 15秒后自动恢复
        });

        // 学生表单提交确认
        document.getElementById('studentUploadForm').addEventListener('submit', function(e) {
            if (!studentFileInput.files.length) {
                e.preventDefault();
                alert('请先选择要上传的学生文件');
                return false;
            }

            const submitBtn = this.querySelector('button[type="submit"]');
            const originalHtml = submitBtn.innerHTML;
            submitBtn.dataset.originalHtml = originalHtml;
            submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>正在生成学生报表...';
            submitBtn.disabled = true;

            // 创建隐藏的iframe用于提交表单，避免页面导航
            const iframe = document.createElement('iframe');
            iframe.name = 'studentDownloadFrame';
            iframe.style.display = 'none';
            document.body.appendChild(iframe);

            // 设置表单target为iframe
            this.target = 'studentDownloadFrame';

            // 监听iframe加载完成（文件下载完成后）
            iframe.onload = function() {
                // 文件下载完成后恢复按钮状态
                setTimeout(function() {
                    submitBtn.innerHTML = originalHtml;
                    submitBtn.disabled = false;
                    // 移除iframe
                    document.body.removeChild(iframe);
                }, 1000);
            };

            // 备用方案：如果iframe的onload未触发（某些浏览器对文件下载的处理不同），使用定时器恢复
            setTimeout(function() {
                if (submitBtn.disabled) {
                    submitBtn.innerHTML = originalHtml;
                    submitBtn.disabled = false;
                    if (iframe.parentNode) {
                        document.body.removeChild(iframe);
                    }
                }
            }, 15000); // 15秒后自动恢复
        });
    </script>
</body>

</html>