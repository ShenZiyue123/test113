<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>每日十分钟，轻松记单词</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- 引入SheetJS库用于解析Excel文件 -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        body {
            background-color: #e0f2e9;
            color: #333;
            min-height: 100vh;
            padding: 20px;
            background-image: linear-gradient(135deg, #c8e6c9, #e0f2e9, #c8e6c9);
            background-size: 400% 400%;
            animation: gradientBG 15s ease infinite;
        }

        @keyframes gradientBG {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }

        .container {
            max-width: 900px;
            margin: 0 auto;
            padding: 20px;
        }

        /* 标题样式 */
        .header {
            display: flex;
            justify-content: center;
            align-items: center;
            margin-bottom: 30px;
            position: relative;
            padding: 20px 0;
            background: rgba(255, 255, 255, 0.7);
            border-radius: 20px;
            box-shadow: 0 8px 20px rgba(46, 125, 50, 0.15);
        }

        .header h1 {
            font-size: 2.8rem;
            color: #2e7d32;
            text-align: center;
            margin: 0 25px;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
            background: linear-gradient(45deg, #2e7d32, #4caf50);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-weight: 700;
            letter-spacing: 1px;
        }

        .decoration {
            font-size: 3.5rem;
            color: #4caf50;
            animation: float 3s ease-in-out infinite;
        }

        .decoration.left {
            transform: rotate(-15deg);
            animation-delay: 0.5s;
        }

        .decoration.right {
            transform: rotate(15deg);
            animation-delay: 1s;
        }

        @keyframes float {
            0% { transform: translateY(0px) rotate(-15deg); }
            50% { transform: translateY(-15px) rotate(-15deg); }
            100% { transform: translateY(0px) rotate(-15deg); }
        }

        /* 卡片样式 */
        .card {
            background-color: rgba(255, 255, 255, 0.92);
            border-radius: 20px;
            box-shadow: 0 12px 25px rgba(0, 0, 0, 0.1);
            padding: 30px;
            margin-bottom: 30px;
            transition: all 0.3s ease;
            border: 1px solid rgba(129, 199, 132, 0.3);
            backdrop-filter: blur(5px);
        }

        .card:hover {
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.15);
            transform: translateY(-8px);
        }

        .card-title {
            color: #2e7d32;
            margin-bottom: 25px;
            font-size: 1.7rem;
            display: flex;
            align-items: center;
            font-weight: 600;
            padding-bottom: 15px;
            border-bottom: 2px solid #e8f5e9;
        }

        .card-title i {
            margin-right: 15px;
            background: #e8f5e9;
            width: 50px;
            height: 50px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 50%;
            font-size: 1.5rem;
        }

        /* 自定义词库区域 */
        .word-input {
            width: 100%;
            min-height: 160px;
            padding: 20px;
            border: 2px dashed #81c784;
            border-radius: 15px;
            font-size: 1.05rem;
            margin-bottom: 25px;
            resize: vertical;
            background-color: #f8fff9;
            transition: all 0.3s ease;
            line-height: 1.6;
        }

        .word-input:focus {
            outline: none;
            border-color: #4caf50;
            box-shadow: 0 0 0 4px rgba(76, 175, 80, 0.2);
        }

        .upload-area {
            display: flex;
            gap: 15px;
            flex-wrap: wrap;
            margin-bottom: 20px;
        }

        .btn {
            padding: 14px 28px;
            border: none;
            border-radius: 50px;
            font-size: 1.05rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.08);
        }

        .btn-primary {
            background: linear-gradient(135deg, #4caf50, #2e7d32);
            color: white;
        }

        .btn-primary:hover {
            background: linear-gradient(135deg, #43a047, #1b5e20);
            transform: translateY(-3px);
            box-shadow: 0 6px 15px rgba(0,0,0,0.12);
        }

        .btn-secondary {
            background: linear-gradient(135deg, #f5f5f5, #e0e0e0);
            color: #333;
            border: 1px solid #e0e0e0;
        }

        .btn-secondary:hover {
            background: linear-gradient(135deg, #e8e8e8, #d5d5d5);
            transform: translateY(-3px);
            box-shadow: 0 6px 15px rgba(0,0,0,0.1);
        }

        .btn-tertiary {
            background: linear-gradient(135deg, #81d4fa, #4fc3f7);
            color: white;
        }

        .btn-tertiary:hover {
            background: linear-gradient(135deg, #4fc3f7, #29b6f6);
            transform: translateY(-3px);
            box-shadow: 0 6px 15px rgba(0,0,0,0.1);
        }

        .btn-danger {
            background: linear-gradient(135deg, #ff7043, #f4511e);
            color: white;
        }

        .btn-danger:hover {
            background: linear-gradient(135deg, #f4511e, #d84315);
            transform: translateY(-3px);
            box-shadow: 0 6px 15px rgba(0,0,0,0.12);
        }

        .btn-warning {
            background: linear-gradient(135deg, #ffca28, #ff9800);
            color: white;
        }

        .btn-warning:hover {
            background: linear-gradient(135deg, #ffb300, #f57c00);
            transform: translateY(-3px);
            box-shadow: 0 6px 15px rgba(0,0,0,0.12);
        }

        .file-input {
            display: none;
        }

        .file-label {
            background: linear-gradient(135deg, #81c784, #66bb6a);
            color: white;
            padding: 14px 28px;
            border-radius: 50px;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 10px;
            transition: all 0.3s ease;
        }

        .file-label:hover {
            background: linear-gradient(135deg, #66bb6a, #4caf50);
            transform: translateY(-3px);
        }

        .hint {
            font-size: 1rem;
            color: #555;
            margin-top: 15px;
            padding: 15px;
            background-color: #f1f8e9;
            border-radius: 12px;
            border-left: 4px solid #4caf50;
            line-height: 1.6;
        }

        .hint i {
            color: #4caf50;
            margin-right: 8px;
        }

        .btn-container {
            display: flex;
            gap: 15px;
            flex-wrap: wrap;
            justify-content: center;
            margin-top: 25px;
        }

        /* 格式示例区域 */
        .format-examples {
            margin: 20px 0;
            padding: 15px;
            background: #f1f8e9;
            border-radius: 12px;
            border-left: 4px solid #4caf50;
        }

        .format-examples h3 {
            margin-bottom: 10px;
            color: #2e7d32;
            display: flex;
            align-items: center;
        }

        .format-examples h3 i {
            margin-right: 10px;
        }

        .examples-container {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-top: 10px;
        }

        .example-item {
            padding: 8px 12px;
            background: white;
            border-radius: 8px;
            font-family: monospace;
            font-size: 0.9rem;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            border: 1px solid #c8e6c9;
        }

        /* Excel格式说明 */
        .excel-format {
            margin: 20px 0;
            padding: 15px;
            background: #e3f2fd;
            border-radius: 12px;
            border-left: 4px solid #2196f3;
        }

        .excel-format h3 {
            margin-bottom: 10px;
            color: #1565c0;
            display: flex;
            align-items: center;
        }

        .excel-format h3 i {
            margin-right: 10px;
        }

        .excel-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .excel-table th, .excel-table td {
            border: 1px solid #bbdefb;
            padding: 10px;
            text-align: center;
        }

        .excel-table th {
            background-color: #bbdefb;
            color: #0d47a1;
            font-weight: 600;
        }

        .excel-table tr:nth-child(even) {
            background-color: #e3f2fd;
        }

        /* 预览区域 */
        .preview-container {
            margin-top: 20px;
            max-height: 200px;
            overflow-y: auto;
            border: 1px solid #c8e6c9;
            border-radius: 12px;
            padding: 15px;
            background: white;
            display: none;
        }

        .preview-title {
            font-weight: bold;
            margin-bottom: 10px;
            color: #2e7d32;
            display: flex;
            align-items: center;
        }

        .preview-title i {
            margin-right: 8px;
        }

        .preview-list {
            list-style: none;
        }

        .preview-item {
            padding: 8px 0;
            border-bottom: 1px solid #e8f5e9;
            display: flex;
        }

        .preview-word {
            font-weight: bold;
            min-width: 120px;
            color: #2e7d32;
        }

        .preview-meaning {
            color: #f57c00;
        }

        /* 单词学习区域 */
        .word-display {
            text-align: center;
            margin: 35px 0;
            padding: 25px;
            min-height: 280px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            background: #f8fff9;
            border-radius: 20px;
            border: 1px solid #e8f5e9;
            box-shadow: inset 0 0 15px rgba(129, 199, 132, 0.1);
        }

        .word {
            font-size: 4rem;
            font-weight: bold;
            color: #2e7d32;
            margin-bottom: 25px;
            letter-spacing: 1px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
            display: flex;
            align-items: center;
            justify-content: center;
            flex-wrap: wrap;
            gap: 20px;
        }

        .speak-icon {
            cursor: pointer;
            color: #4caf50;
            font-size: 2.5rem;
            transition: all 0.3s ease;
            background: #e8f5e9;
            width: 60px;
            height: 60px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 50%;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }

        .speak-icon:hover {
            background: #4caf50;
            color: white;
            transform: scale(1.1);
        }

        .speak-icon.speaking {
            animation: pulse 1s infinite;
            background: #4caf50;
            color: white;
        }

        .meaning {
            font-size: 2.2rem;
            color: #f57c00;
            margin-bottom: 20px;
            min-height: 55px;
            opacity: 0;
            transition: opacity 0.5s ease;
            font-weight: 600;
            display: flex;
            align-items: center;
            justify-content: center;
            flex-wrap: wrap;
            gap: 20px;
        }

        .meaning.show {
            opacity: 1;
        }

        .example {
            font-size: 1.4rem;
            color: #555;
            font-style: italic;
            margin-bottom: 35px;
            max-width: 700px;
            line-height: 1.7;
            padding: 15px 25px;
            background: #f1f8e9;
            border-radius: 15px;
            border-left: 4px solid #4caf50;
            display: flex;
            align-items: center;
            justify-content: center;
            flex-wrap: wrap;
            gap: 20px;
        }

        .controls {
            display: flex;
            justify-content: center;
            flex-wrap: wrap;
            gap: 20px;
            margin-top: 30px;
        }

        .btn-word {
            min-width: 180px;
            padding: 16px 10px;
        }

        .progress-container {
            margin-top: 35px;
            width: 100%;
            background-color: #e0e0e0;
            border-radius: 10px;
            overflow: hidden;
            height: 14px;
        }

        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #81c784, #4caf50);
            width: 0%;
            transition: width 0.8s ease;
        }

        .stats {
            display: flex;
            justify-content: space-around;
            margin-top: 25px;
            color: #555;
            font-size: 1.1rem;
            font-weight: 500;
            background: #f1f8e9;
            padding: 15px;
            border-radius: 15px;
        }

        .stats div {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .hidden {
            display: none;
        }

        /* 模态框样式 */
        .modal {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.7);
            z-index: 1000;
            align-items: center;
            justify-content: center;
            opacity: 0;
            transition: opacity 0.3s ease;
        }

        .modal.show {
            display: flex;
            opacity: 1;
        }

        .modal-content {
            background: white;
            width: 90%;
            max-width: 700px;
            max-height: 80vh;
            border-radius: 20px;
            overflow: hidden;
            box-shadow: 0 25px 50px rgba(0,0,0,0.25);
            transform: translateY(-30px);
            transition: transform 0.4s ease;
        }

        .modal.show .modal-content {
            transform: translateY(0);
        }

        .modal-header {
            background: linear-gradient(135deg, #4caf50, #2e7d32);
            color: white;
            padding: 25px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .modal-header h2 {
            font-size: 1.8rem;
            font-weight: 600;
        }

        .close-btn {
            background: none;
            border: none;
            color: white;
            font-size: 1.8rem;
            cursor: pointer;
            transition: transform 0.3s ease;
        }

        .close-btn:hover {
            transform: rotate(90deg);
        }

        .modal-body {
            padding: 25px;
            overflow-y: auto;
            max-height: 60vh;
        }

        .error-word-list {
            list-style: none;
            padding: 0;
        }

        .error-word-item {
            padding: 15px 20px;
            border-bottom: 1px solid #eee;
            display: flex;
            align-items: center;
            gap: 15px;
            transition: background 0.3s ease;
            font-size: 1.4rem;
        }

        .error-word-item:hover {
            background: #f9f9f9;
        }

        .error-word {
            font-weight: bold;
            color: #2e7d32;
            min-width: 200px;
            text-align: left;
        }

        .error-meaning {
            color: #f57c00;
            flex: 1;
        }

        .empty-message {
            text-align: center;
            padding: 40px;
            color: #888;
            font-size: 1.2rem;
        }

        .modal-footer {
            padding: 20px;
            background: #f9f9f9;
            display: flex;
            justify-content: flex-end;
            gap: 15px;
        }

        /* 响应式设计 */
        @media (max-width: 768px) {
            .header h1 {
                font-size: 2rem;
            }
            
            .decoration {
                font-size: 2.5rem;
            }
            
            .word {
                font-size: 2.8rem;
            }
            
            .btn {
                padding: 12px 20px;
                font-size: 0.95rem;
            }
            
            .controls {
                flex-direction: column;
                align-items: center;
            }
            
            .btn-word {
                width: 100%;
                max-width: 320px;
            }
            
            .upload-area {
                flex-direction: column;
                align-items: center;
            }
            
            .file-label, .btn {
                width: 100%;
                justify-content: center;
            }
            
            .modal-footer {
                flex-wrap: wrap;
            }
            
            .modal-footer .btn {
                flex: 1;
                min-width: 120px;
            }
            
            .error-word-item {
                flex-direction: column;
                align-items: flex-start;
                gap: 8px;
            }
            
            .error-word {
                min-width: auto;
            }
        }

        /* 动画效果 */
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .fade-in {
            animation: fadeIn 0.8s ease forwards;
        }
        
        .pulse {
            animation: pulse 2s infinite;
        }
        
        @keyframes pulse {
            0% { transform: scale(1); }
            50% { transform: scale(1.05); }
            100% { transform: scale(1); }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- 标题区域 -->
        <div class="header">
            <div class="decoration left">📚</div>
            <h1>每日十分钟，轻松记单词</h1>
            <div class="decoration right">✏️</div>
        </div>

        <!-- 自定义词库区域 -->
        <div class="card">
            <div class="card-title">
                <i class="fas fa-book"></i>
                自定义词库
            </div>
            
            <textarea class="word-input" id="wordInput" placeholder="请输入单词和词义，支持多种格式（例如：apple 苹果、apple:苹果、apple-苹果等）。每行一个单词。如果不输入，将使用默认词库。">apple 苹果
book:书
cat-猫
dog——狗
friend 朋友
knowledge 知识
nature 自然
ocean 海洋
river 河流
teacher 老师
look forward to 期待
artificial intelligence 人工智能</textarea>
            
            <!-- 格式示例 -->
            <div class="format-examples">
                <h3><i class="fas fa-lightbulb"></i> 支持的格式示例</h3>
                <div class="examples-container">
                    <div class="example-item">apple 苹果</div>
                    <div class="example-item">book:书</div>
                    <div class="example-item">cat-猫</div>
                    <div class="example-item">dog——狗</div>
                    <div class="example-item">friend,朋友</div>
                    <div class="example-item">hello;你好</div>
                    <div class="example-item">look forward to 期待</div>
                    <div class="example-item">artificial intelligence:人工智能</div>
                </div>
            </div>
            
            <!-- Excel格式说明 -->
            <div class="excel-format">
                <h3><i class="fas fa-file-excel"></i> Excel格式说明</h3>
                <p>上传Excel文件时，请确保文件格式如下：</p>
                <table class="excel-table">
                    <thead>
                        <tr>
                            <th>A列 (单词/短语)</th>
                            <th>B列 (中文释义)</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>apple</td>
                            <td>苹果</td>
                        </tr>
                        <tr>
                            <td>look forward to</td>
                            <td>期待，盼望</td>
                        </tr>
                        <tr>
                            <td>artificial intelligence</td>
                            <td>人工智能</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <!-- 预览区域 -->
            <div class="preview-container" id="previewContainer">
                <div class="preview-title"><i class="fas fa-eye"></i> 导入预览</div>
                <ul class="preview-list" id="previewList"></ul>
            </div>
            
            <div class="upload-area">
                <input type="file" id="txtFileInput" class="file-input" accept=".txt">
                <label for="txtFileInput" class="file-label">
                    <i class="fas fa-upload"></i> 上传TXT词库
                </label>
                
                <input type="file" id="excelFileInput" class="file-input" accept=".xlsx,.xls,.csv">
                <label for="excelFileInput" class="file-label">
                    <i class="fas fa-file-excel"></i> 上传Excel词库
                </label>
                
                <button class="btn btn-secondary" id="previewBtn">
                    <i class="fas fa-eye"></i> 预览导入
                </button>
                
                <button class="btn btn-tertiary" id="exportErrorBtn">
                    <i class="fas fa-file-export"></i> 导出易错词汇本
                </button>
                
                <button class="btn btn-tertiary" id="viewErrorBtn">
                    <i class="fas fa-history"></i> 查看历史易错词汇
                </button>
                
                <button class="btn btn-warning" id="clearSavedBtn">
                    <i class="fas fa-trash"></i> 清空保存的词库
                </button>
            </div>
            
            <div class="hint">
                <i class="fas fa-info-circle"></i> 提示：支持多种分隔符（空格、逗号、冒号、分号、破折号等）。每行一个单词。同一个单词点击三次"还不熟"会自动加入易错词汇本。
            </div>
            
            <div class="btn-container">
                <button class="btn btn-primary pulse" id="startBtn">
                    <i class="fas fa-play"></i> 开始学习
                </button>
            </div>
        </div>

        <!-- 单词学习区域 -->
        <div class="card hidden" id="wordCard">
            <div class="word-display">
                <div class="word" id="currentWord">welcome <i class="fas fa-volume-up speak-icon" id="wordSpeakIcon"></i></div>
                <div class="meaning" id="wordMeaning">欢迎 <i class="fas fa-volume-up speak-icon" id="meaningSpeakIcon"></i></div>
                <div class="example" id="wordExample">Welcome to our vocabulary learning program! <i class="fas fa-volume-up speak-icon" id="exampleSpeakIcon"></i></div>
            </div>
            
            <div class="controls">
                <button class="btn btn-primary btn-word" id="showMeaningBtn">
                    <i class="fas fa-eye"></i> 显示词义
                </button>
                <button class="btn btn-primary btn-word" id="masteredBtn">
                    <i class="fas fa-check-circle"></i> 已掌握
                </button>
                <button class="btn btn-secondary btn-word" id="notFamiliarBtn">
                    <i class="fas fa-redo"></i> 还不熟
                </button>
                <button class="btn btn-secondary btn-word" id="addToErrorBtn">
                    <i class="fas fa-bookmark"></i> 加入易错词汇本
                </button>
                <button class="btn btn-danger btn-word" id="endGameBtn">
                    <i class="fas fa-stop"></i> 结束学习
                </button>
            </div>
            
            <div class="progress-container">
                <div class="progress-bar" id="progressBar"></div>
            </div>
            
            <div class="stats">
                <div><i class="fas fa-list"></i> 剩余: <span id="remainingCount">0</span></div>
                <div><i class="fas fa-bookmark"></i> 易错: <span id="errorCount">0</span></div>
                <div><i class="fas fa-check"></i> 已掌握: <span id="masteredCount">0</span></div>
            </div>
        </div>
    </div>
    
    <!-- 历史易错词汇模态框 -->
    <div class="modal" id="errorModal">
        <div class="modal-content">
            <div class="modal-header">
                <h2><i class="fas fa-history"></i> 历史易错词汇</h2>
                <button class="close-btn">&times;</button>
            </div>
            <div class="modal-body">
                <ul class="error-word-list" id="errorWordList">
                    <!-- 动态填充内容 -->
                </ul>
            </div>
            <div class="modal-footer">
                <button class="btn btn-primary" id="copyErrorBtn">
                    <i class="fas fa-copy"></i> 复制所有内容
                </button>
                <button class="btn btn-warning" id="clearErrorBtn">
                    <i class="fas fa-trash"></i> 清空
                </button>
                <button class="btn btn-danger" id="closeModalBtn">
                    <i class="fas fa-times"></i> 关闭
                </button>
            </div>
        </div>
    </div>

    <script>
        // 默认词库（初中高频词汇）
        const defaultWords = [
            { word: "apple", meaning: "苹果", example: "I eat an apple every day." },
            { word: "book", meaning: "书", example: "This book is very interesting." },
            { word: "cat", meaning: "猫", example: "The cat is sleeping on the sofa." },
            { word: "dog", meaning: "狗", example: "My dog likes to play in the park." },
            { word: "elephant", meaning: "大象", example: "Elephants have long trunks." },
            { word: "friend", meaning: "朋友", example: "A good friend is always there for you." },
            { word: "garden", meaning: "花园", example: "We grow vegetables in our garden." },
            { word: "happy", meaning: "快乐的", example: "The children are happy on the playground." },
            { word: "internet", meaning: "互联网", example: "We use the internet to find information." },
            { word: "journey", meaning: "旅行", example: "Our journey to the mountains was exciting." },
            { word: "knowledge", meaning: "知识", example: "Reading books increases your knowledge." },
            { word: "library", meaning: "图书馆", example: "I borrow books from the library every week." },
            { word: "mountain", meaning: "山", example: "We climbed the mountain yesterday." },
            { word: "nature", meaning: "自然", example: "We should protect nature." },
            { word: "ocean", meaning: "海洋", example: "The ocean is full of amazing creatures." },
            { word: "pencil", meaning: "铅笔", example: "I write with a blue pencil." },
            { word: "question", meaning: "问题", example: "Do you have any questions?" },
            { word: "river", meaning: "河流", example: "The river flows into the sea." },
            { word: "sun", meaning: "太阳", example: "The sun rises in the east." },
            { word: "teacher", meaning: "老师", example: "Our teacher explains things clearly." }
        ];

        // 全局变量
        let wordList = [];
        let currentWordIndex = -1;
        let masteredWords = [];
        let errorProneWords = []; // 只存储单词和词义
        let notFamiliarCounts = {};
        let remainingWords = [];
        let speechSynthesis = window.speechSynthesis;

        // DOM元素
        const wordInput = document.getElementById('wordInput');
        const txtFileInput = document.getElementById('txtFileInput');
        const excelFileInput = document.getElementById('excelFileInput');
        const previewBtn = document.getElementById('previewBtn');
        const previewContainer = document.getElementById('previewContainer');
        const previewList = document.getElementById('previewList');
        const startBtn = document.getElementById('startBtn');
        const exportErrorBtn = document.getElementById('exportErrorBtn');
        const viewErrorBtn = document.getElementById('viewErrorBtn');
        const clearSavedBtn = document.getElementById('clearSavedBtn');
        const wordCard = document.getElementById('wordCard');
        const currentWordEl = document.getElementById('currentWord');
        const wordMeaningEl = document.getElementById('wordMeaning');
        const wordExampleEl = document.getElementById('wordExample');
        const showMeaningBtn = document.getElementById('showMeaningBtn');
        const masteredBtn = document.getElementById('masteredBtn');
        const notFamiliarBtn = document.getElementById('notFamiliarBtn');
        const addToErrorBtn = document.getElementById('addToErrorBtn');
        const endGameBtn = document.getElementById('endGameBtn');
        const progressBar = document.getElementById('progressBar');
        const remainingCountEl = document.getElementById('remainingCount');
        const errorCountEl = document.getElementById('errorCount');
        const masteredCountEl = document.getElementById('masteredCount');
        const errorModal = document.getElementById('errorModal');
        const errorWordList = document.getElementById('errorWordList');
        const copyErrorBtn = document.getElementById('copyErrorBtn');
        const clearErrorBtn = document.getElementById('clearErrorBtn');
        const closeModalBtn = document.getElementById('closeModalBtn');
        const closeBtn = document.querySelector('.close-btn');
        
        // 语音朗读相关元素
        const wordSpeakIcon = document.getElementById('wordSpeakIcon');
        const meaningSpeakIcon = document.getElementById('meaningSpeakIcon');
        const exampleSpeakIcon = document.getElementById('exampleSpeakIcon');

        // 初始化
        document.addEventListener('DOMContentLoaded', () => {
            // 尝试从localStorage加载保存的词库和易错词汇
            loadSavedWordList();
            loadErrorProneWords();
            updateStats();
            
            // 绑定事件
            wordInput.addEventListener('input', debounce(handleInputChange, 500));
            previewBtn.addEventListener('click', previewWords);
            startBtn.addEventListener('click', startLearning);
            txtFileInput.addEventListener('change', handleTxtFileUpload);
            excelFileInput.addEventListener('change', handleExcelFileUpload);
            exportErrorBtn.addEventListener('click', exportErrorWords);
            viewErrorBtn.addEventListener('click', showErrorWords);
            clearSavedBtn.addEventListener('click', clearSavedWordList);
            showMeaningBtn.addEventListener('click', showMeaning);
            masteredBtn.addEventListener('click', markAsMastered);
            notFamiliarBtn.addEventListener('click', markAsNotFamiliar);
            addToErrorBtn.addEventListener('click', addToErrorProne);
            endGameBtn.addEventListener('click', endGame);
            copyErrorBtn.addEventListener('click', copyErrorWords);
            clearErrorBtn.addEventListener('click', clearErrorWords);
            closeModalBtn.addEventListener('click', closeModal);
            closeBtn.addEventListener('click', closeModal);
            
            // 语音朗读事件
            wordSpeakIcon.addEventListener('click', () => speakWord(currentWordEl.textContent.replace('🔊', '').trim()));
            meaningSpeakIcon.addEventListener('click', () => speakWord(wordMeaningEl.textContent.replace('🔊', '').trim()));
            exampleSpeakIcon.addEventListener('click', () => speakWord(wordExampleEl.textContent.replace('🔊', '').trim()));
            
            // 点击模态框外部关闭
            errorModal.addEventListener('click', (e) => {
                if (e.target === errorModal) closeModal();
            });
        });

        // 语音朗读函数
        function speakWord(text) {
            // 如果正在朗读，先停止
            speechSynthesis.cancel();
            
            // 检查浏览器是否支持语音合成
            if (!speechSynthesis) {
                alert('您的浏览器不支持语音朗读功能，请使用Chrome、Edge或Safari等现代浏览器。');
                return;
            }
            
            // 创建语音合成实例
            const utterance = new SpeechSynthesisUtterance(text);
            
            // 设置语言为英语
            utterance.lang = 'en-US';
            
            // 设置语速和音调
            utterance.rate = 0.9;
            utterance.pitch = 1.0;
            
            // 朗读开始时的回调
            utterance.onstart = function() {
                // 为当前朗读的图标添加动画效果
                document.querySelectorAll('.speak-icon').forEach(icon => {
                    icon.classList.remove('speaking');
                });
                event.target.classList.add('speaking');
            };
            
            // 朗读结束时的回调
            utterance.onend = function() {
                document.querySelectorAll('.speak-icon').forEach(icon => {
                    icon.classList.remove('speaking');
                });
            };
            
            // 朗读错误的回调
            utterance.onerror = function(event) {
                console.error('语音朗读错误:', event);
                document.querySelectorAll('.speak-icon').forEach(icon => {
                    icon.classList.remove('speaking');
                });
                alert('语音朗读出现错误，请稍后再试。');
            };
            
            // 开始朗读
            speechSynthesis.speak(utterance);
        }

        // 防抖函数
        function debounce(func, wait) {
            let timeout;
            return function executedFunction(...args) {
                const later = () => {
                    clearTimeout(timeout);
                    func(...args);
                };
                clearTimeout(timeout);
                timeout = setTimeout(later, wait);
            };
        }

        // 处理输入变化
        function handleInputChange() {
            const userInput = wordInput.value.trim();
            if (userInput) {
                previewWords();
                // 保存词库到localStorage
                saveWordList();
            } else {
                previewContainer.style.display = 'none';
                // 如果输入为空，也保存到localStorage
                saveWordList();
            }
        }

        // 预览单词
        function previewWords() {
            const userInput = wordInput.value.trim();
            if (!userInput) {
                previewContainer.style.display = 'none';
                return;
            }
            
            const parsedWords = parseWordInput(userInput);
            renderPreview(parsedWords);
        }

        // 渲染预览
        function renderPreview(words) {
            previewList.innerHTML = '';
            
            if (words.length === 0) {
                previewList.innerHTML = '<li>未识别到有效单词，请检查格式</li>';
                previewContainer.style.display = 'block';
                return;
            }
            
            words.forEach(word => {
                const li = document.createElement('li');
                li.className = 'preview-item';
                li.innerHTML = `
                    <span class="preview-word">${word.word}</span>
                    <span class="preview-meaning">${word.meaning}</span>
                `;
                previewList.appendChild(li);
            });
            
            previewContainer.style.display = 'block';
        }

        // 开始学习
        function startLearning() {
            // 获取用户输入的单词
            const userInput = wordInput.value.trim();
            
            if (userInput) {
                // 解析用户输入的单词
                wordList = parseWordInput(userInput);
            } else {
                // 使用默认词库
                wordList = [...defaultWords];
            }
            
            if (wordList.length === 0) {
                alert('词库为空，请添加单词或使用默认词库！');
                return;
            }
            
            // 初始化学习状态
            currentWordIndex = -1;
            masteredWords = [];
            notFamiliarCounts = {};
            remainingWords = [...wordList];
            
            // 打乱单词顺序
            shuffleArray(remainingWords);
            
            // 显示单词卡片
            wordCard.classList.remove('hidden');
            
            // 滚动到单词卡片
            wordCard.scrollIntoView({ behavior: 'smooth' });
            
            // 显示第一个单词
            showNextWord();
        }

        // 解析用户输入的单词（改进版）- 修复短语识别问题
        function parseWordInput(input) {
            const lines = input.split('\n');
            const words = [];
            
            // 非空格分隔符的正则表达式
            const nonSpaceSeparators = /[:：;；,，—-]+/;
            
            for (const line of lines) {
                const trimmedLine = line.trim();
                if (trimmedLine) {
                    // 跳过注释行（以#开头的行）
                    if (trimmedLine.startsWith('#')) continue;
                    
                    let word = '';
                    let meaning = '';
                    
                    // 首先检查是否有非空格分隔符
                    const nonSpaceMatch = trimmedLine.match(nonSpaceSeparators);
                    if (nonSpaceMatch) {
                        const separatorIndex = trimmedLine.search(nonSpaceSeparators);
                        word = trimmedLine.substring(0, separatorIndex).trim();
                        meaning = trimmedLine.substring(separatorIndex + nonSpaceMatch[0].length).trim();
                    } else {
                        // 如果没有非空格分隔符，使用第一个空格作为分隔符
                        const firstSpaceIndex = trimmedLine.indexOf(' ');
                        if (firstSpaceIndex !== -1) {
                            word = trimmedLine.substring(0, firstSpaceIndex).trim();
                            meaning = trimmedLine.substring(firstSpaceIndex + 1).trim();
                        } else {
                            // 如果没有分隔符，整行作为单词，释义为空
                            word = trimmedLine;
                            meaning = "（未提供释义）";
                        }
                    }
                    
                    if (word) {
                        // 查找默认词库中的例句
                        const defaultWord = defaultWords.find(w => w.word === word);
                        const example = defaultWord ? defaultWord.example : `Try to use the word "${word}" in a sentence.`;
                        
                        words.push({ word, meaning, example });
                    }
                }
            }
            
            return words;
        }

        // 处理TXT文件上传
        function handleTxtFileUpload(event) {
            const file = event.target.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = function(e) {
                wordInput.value = e.target.result;
                // 保存词库到localStorage
                saveWordList();
                previewWords();
            };
            reader.readAsText(file);
            
            // 重置文件输入，允许再次选择同一个文件
            event.target.value = '';
        }

        // 处理Excel文件上传
        function handleExcelFileUpload(event) {
            const file = event.target.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // 获取第一个工作表
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    
                    // 将工作表转换为JSON
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: ['word', 'meaning'] });
                    
                    // 过滤掉标题行（如果有）
                    const words = jsonData.filter(item => 
                        item.word && item.meaning && 
                        item.word.toString().toLowerCase() !== 'word' && 
                        item.meaning.toString().toLowerCase() !== 'meaning'
                    ).map(item => {
                        const word = item.word.toString().trim();
                        const meaning = item.meaning.toString().trim();
                        
                        // 查找默认词库中的例句
                        const defaultWord = defaultWords.find(w => w.word === word);
                        const example = defaultWord ? defaultWord.example : `Try to use the word "${word}" in a sentence.`;
                        
                        return { word, meaning, example };
                    });
                    
                    if (words.length === 0) {
                        alert('Excel文件中未找到有效数据！请确保第一列是单词，第二列是释义。');
                        return;
                    }
                    
                    // 生成文本格式的内容
                    const textContent = words.map(item => `${item.word} ${item.meaning}`).join('\n');
                    wordInput.value = textContent;
                    
                    // 保存词库到localStorage
                    saveWordList();
                    
                    // 显示预览
                    renderPreview(words);
                    previewContainer.style.display = 'block';
                    
                    alert(`成功导入 ${words.length} 个单词和短语！`);
                } catch (error) {
                    console.error('Excel解析错误:', error);
                    alert('Excel文件解析失败，请检查文件格式！');
                }
            };
            
            reader.onerror = function() {
                alert('文件读取失败！');
            };
            
            reader.readAsArrayBuffer(file);
            
            // 重置文件输入，允许再次选择同一个文件
            event.target.value = '';
        }

        // 显示下一个单词
        function showNextWord() {
            // 检查是否还有单词
            if (remainingWords.length === 0) {
                endGame();
                return;
            }
            
            // 更新进度条
            const progress = ((wordList.length - remainingWords.length) / wordList.length) * 100;
            progressBar.style.width = `${progress}%`;
            
            // 更新统计信息
            updateStats();
            
            // 获取下一个单词
            currentWordIndex = (currentWordIndex + 1) % remainingWords.length;
            const wordObj = remainingWords[currentWordIndex];
            
            // 显示单词
            currentWordEl.innerHTML = `${wordObj.word} <i class="fas fa-volume-up speak-icon" id="wordSpeakIcon"></i>`;
            wordMeaningEl.innerHTML = `${wordObj.meaning} <i class="fas fa-volume-up speak-icon" id="meaningSpeakIcon"></i>`;
            wordMeaningEl.classList.remove('show');
            wordExampleEl.innerHTML = `${wordObj.example} <i class="fas fa-volume-up speak-icon" id="exampleSpeakIcon"></i>`;
            
            // 重新绑定语音朗读事件
            document.getElementById('wordSpeakIcon').addEventListener('click', () => speakWord(wordObj.word));
            document.getElementById('meaningSpeakIcon').addEventListener('click', () => speakWord(wordObj.meaning));
            document.getElementById('exampleSpeakIcon').addEventListener('click', () => speakWord(wordObj.example));
            
            // 添加淡入动画
            currentWordEl.classList.remove('fade-in');
            void currentWordEl.offsetWidth; // 触发重绘
            currentWordEl.classList.add('fade-in');
        }

        // 显示词义
        function showMeaning() {
            wordMeaningEl.classList.add('show');
        }

        // 标记为已掌握
        function markAsMastered() {
            const currentWord = remainingWords[currentWordIndex];
            masteredWords.push(currentWord);
            removeCurrentWord();
            showNextWord();
        }

        // 标记为不熟悉
        function markAsNotFamiliar() {
            const currentWord = remainingWords[currentWordIndex];
            const word = currentWord.word;
            
            // 更新不熟悉次数
            notFamiliarCounts[word] = (notFamiliarCounts[word] || 0) + 1;
            
            // 如果不熟悉次数达到3次，加入易错词汇本
            if (notFamiliarCounts[word] >= 3) {
                if (!errorProneWords.some(w => w.word === word)) {
                    // 只保存单词和词义，不保存例句
                    errorProneWords.push({ 
                        word: currentWord.word, 
                        meaning: currentWord.meaning 
                    });
                    saveErrorProneWords();
                    alert(`单词 "${word}" 已自动加入易错词汇本！`);
                }
            }
            
            // 移动到下一个单词
            showNextWord();
        }

        // 添加到易错词汇本
        function addToErrorProne() {
            const currentWord = remainingWords[currentWordIndex];
            const word = currentWord.word;
            
            if (!errorProneWords.some(w => w.word === word)) {
                // 只保存单词和词义，不保存例句
                errorProneWords.push({ 
                    word: currentWord.word, 
                    meaning: currentWord.meaning 
                });
                saveErrorProneWords();
                alert(`单词 "${word}" 已加入易错词汇本！`);
            } else {
                alert(`单词 "${word}" 已经在易错词汇本中！`);
            }
            
            removeCurrentWord();
            showNextWord();
        }

        // 结束学习
        function endGame() {
            // 停止语音朗读
            speechSynthesis.cancel();
            
            wordCard.classList.add('hidden');
            alert(`学习结束！\n已掌握单词: ${masteredWords.length}\n易错单词: ${errorProneWords.length}`);
        }

        // 导出易错词汇本
        function exportErrorWords() {
            if (errorProneWords.length === 0) {
                alert('易错词汇本为空！');
                return;
            }
            
            // 创建文本内容（只包含单词和词义）
            let content = '';
            for (const word of errorProneWords) {
                content += `${word.word} ${word.meaning}\n`;
            }
            
            // 创建下载链接
            const blob = new Blob([content], { type: 'text/plain' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = '易错词汇本.txt';
            document.body.appendChild(a);
            a.click();
            
            // 清理
            setTimeout(() => {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 0);
        }
        
        // 显示历史易错词汇
        function showErrorWords() {
            // 填充词汇列表
            renderErrorWords();
            
            // 显示模态框
            errorModal.classList.add('show');
            document.body.style.overflow = 'hidden';
        }
        
        // 渲染易错词汇列表（只显示单词和词义）
        function renderErrorWords() {
            errorWordList.innerHTML = '';
            
            if (errorProneWords.length === 0) {
                errorWordList.innerHTML = `
                    <div class="empty-message">
                        <i class="fas fa-inbox" style="font-size: 3rem; margin-bottom: 15px;"></i>
                        <p>当前没有历史易错词汇</p>
                        <p style="font-size: 0.9rem; margin-top: 10px;">开始学习并标记一些易错词汇吧！</p>
                    </div>
                `;
            } else {
                errorProneWords.forEach(word => {
                    const li = document.createElement('li');
                    li.className = 'error-word-item';
                    li.innerHTML = `
                        <span class="error-word">${word.word}</span>
                        <span class="error-meaning">${word.meaning}</span>
                    `;
                    errorWordList.appendChild(li);
                });
            }
        }
        
        // 复制易错词汇（只复制单词和词义）
        function copyErrorWords() {
            if (errorProneWords.length === 0) {
                alert('没有可复制的内容！');
                return;
            }
            
            // 创建文本内容（只包含单词和词义）
            let content = '';
            for (const word of errorProneWords) {
                content += `${word.word} ${word.meaning}\n`;
            }
            
            // 复制到剪贴板
            navigator.clipboard.writeText(content)
                .then(() => {
                    alert('易错词汇已复制到剪贴板！');
                })
                .catch(err => {
                    console.error('复制失败: ', err);
                    alert('复制失败，请手动复制内容！');
                });
        }
        
        // 清空历史易错词汇
        function clearErrorWords() {
            if (errorProneWords.length === 0) {
                alert('当前没有可清空的易错词汇！');
                return;
            }
            
            if (confirm('确定要清空所有历史易错词汇吗？此操作不可撤销！')) {
                // 清空易错词汇
                errorProneWords = [];
                saveErrorProneWords();
                
                // 更新统计信息
                updateStats();
                
                // 重新渲染词汇列表
                renderErrorWords();
                
                alert('历史易错词汇已成功清空！');
            }
        }
        
        // 关闭模态框
        function closeModal() {
            errorModal.classList.remove('show');
            document.body.style.overflow = 'auto';
        }

        // 辅助函数：从剩余单词中移除当前单词
        function removeCurrentWord() {
            remainingWords.splice(currentWordIndex, 1);
            currentWordIndex = Math.max(0, currentWordIndex - 1);
        }

        // 辅助函数：打乱数组顺序
        function shuffleArray(array) {
            for (let i = array.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                [array[i], array[j]] = [array[j], array[i]];
            }
        }

        // 辅助函数：更新统计信息
        function updateStats() {
            remainingCountEl.textContent = remainingWords.length;
            errorCountEl.textContent = errorProneWords.length;
            masteredCountEl.textContent = masteredWords.length;
        }

        // 辅助函数：保存易错词汇到localStorage
        function saveErrorProneWords() {
            localStorage.setItem('errorProneWords', JSON.stringify(errorProneWords));
            updateStats();
        }

        // 辅助函数：从localStorage加载易错词汇
        function loadErrorProneWords() {
            const savedWords = localStorage.getItem('errorProneWords');
            if (savedWords) {
                try {
                    errorProneWords = JSON.parse(savedWords);
                    
                    // 转换旧格式（如果存在）
                    if (errorProneWords.length > 0 && errorProneWords[0].example) {
                        errorProneWords = errorProneWords.map(item => ({
                            word: item.word,
                            meaning: item.meaning
                        }));
                        saveErrorProneWords(); // 保存转换后的格式
                    }
                } catch (e) {
                    errorProneWords = [];
                }
            }
        }

        // 辅助函数：保存词库到localStorage
        function saveWordList() {
            localStorage.setItem('savedWordList', wordInput.value);
        }

        // 辅助函数：从localStorage加载保存的词库
        function loadSavedWordList() {
            const savedWordList = localStorage.getItem('savedWordList');
            if (savedWordList) {
                wordInput.value = savedWordList;
                // 显示预览
                previewWords();
            }
        }

        // 辅助函数：清空保存的词库
        function clearSavedWordList() {
            if (confirm('确定要清空保存的词库吗？此操作不可撤销！')) {
                localStorage.removeItem('savedWordList');
                wordInput.value = '';
                previewContainer.style.display = 'none';
                alert('保存的词库已成功清空！');
            }
        }
    </script>
</body>
</html>
