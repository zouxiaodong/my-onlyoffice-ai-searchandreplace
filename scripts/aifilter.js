(function (window, undefined) {
    var Api;

    var pluginConfig = {
        backendUrl: 'http://192.168.4.43:8001/api/fill-document',
        licenseImageUrl: 'https://n.sinaimg.cn/news/1_img/upload/0680838e/317/w2048h1469/20250809/4d69-65647e059977698f6052f6064c507dc2.jpg'
    };

    function updateStatus(message, isError = false) {
        var statusElement = document.getElementById('status-message');
        if (statusElement) {
            statusElement.textContent = message;
            statusElement.classList.toggle('error', isError);
        }
    }

    // 一键填充处理函数（只做文本替换）
    function onFillButtonClick() {
        var button = document.getElementById('fill-button');
        if (button) button.disabled = true;
        updateStatus('正在读取文档内容...');

        // 第一次callCommand：获取文档内容并返回
        window.Asc.plugin.callCommand(function () {
            var oDocument = Api.GetDocument();
            if (!oDocument) return '';

            var text = oDocument.GetText() || '';
            console.log('callCommand中获取文档内容:', text);

            // 返回文档文本给回调函数
            return text;
        }, false, false, function (returnedText) {
            console.log('回调收到文档内容:', returnedText);

            if (!returnedText || returnedText.length === 0) {
                updateStatus('未获取到文档内容', true);
                if (button) button.disabled = false;
                return;
            }

            updateStatus('正在发送到 AI 服务...');

            // 发送给AI后端
            fetch(pluginConfig.backendUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ documentContent: returnedText })
            })
                .then(resp => {
                    if (!resp.ok) throw new Error('HTTP状态 ' + resp.status);
                    return resp.json();
                })
                .then(data => {
                    if (!data.success) throw new Error(data.error || 'AI服务失败');

                    var replacements = data.replacements || [];
                    if (replacements.length === 0) {
                        updateStatus('未检测到可替换内容');
                        if (button) button.disabled = false;
                        return;
                    }

                    // 第二次callCommand：执行替换操作
                    Asc.scope.replacements = JSON.stringify(replacements);

                    window.Asc.plugin.callCommand(function () {
                        var reps = JSON.parse(Asc.scope.replacements);
                        var oDoc = Api.GetDocument();
                        if (!oDoc) return;

                        reps.forEach(function (replacement) {
                            var target = replacement.target;
                            var value = replacement.value;
                            if (target && value !== undefined && value !== null) {
                                console.log('替换:', target, '->', value);
                                oDoc.SearchAndReplace({
                                    searchString: target,
                                    replaceString: value
                                });
                            }
                        });
                    }, false, false, function () {
                        updateStatus('文档填充完成！');
                        console.log('替换完成');
                        if (button) button.disabled = false;
                    });
                })
                .catch(err => {
                    updateStatus('请求失败: ' + err.message, true);
                    if (button) button.disabled = false;
                    console.error(err);
                });
        });
    }

    // 插入营业执照图片（单独功能）
    function onInsertImageClick() {
        var button = document.getElementById('insert-license-button');
        if (button) button.disabled = true;
        updateStatus('正在插入营业执照图片...');

        // 确保图片URL是有效的
        if (!pluginConfig.licenseImageUrl) {
            updateStatus('图片URL无效', true);
            return;
        }

        // 先传图片URL到沙箱变量
        Asc.scope.licenseImageUrl = pluginConfig.licenseImageUrl;

        window.Asc.plugin.callCommand(function () {
            var oDocument = Api.GetDocument();
            if (!oDocument) return;

            // 获取当前光标位置
            var oSelection = oDocument.GetSelection();
            if (!oSelection) {
                updateStatus('无法获取光标位置', true);
                return;
            }

            // 创建图像对象，使用 CreateImage
            var oDrawing = Api.CreateImage(Asc.scope.licenseImageUrl, 5000000, 6000000);  // 设置合适的图片尺寸

            // 设置图片的属性（如位置）
            oDrawing.SetWrappingStyle("inFront");  // 图片覆盖文本
            oDrawing.SetHorPosition("cursor");   // 设置水平位置为光标位置
            oDrawing.SetVerPosition("cursor");   // 设置垂直位置为光标位置

            // 将图像插入光标所在的位置
            oSelection.InsertDrawing(oDrawing);

            console.log('营业执照图片插入成功');
        }, false, false, function () {
            updateStatus('营业执照图片插入完成！');
            console.log('营业执照图片插入成功');
            if (button) button.disabled = false;
        });
    }



    window.Asc.plugin.init = function (api) {
        Api = api;

        var fillButton = document.getElementById('fill-button');
        if (fillButton) {
            fillButton.addEventListener('click', onFillButtonClick);
        } else {
            console.error('找不到 fill-button');
        }

        var insertImageButton = document.getElementById('insert-license-button');
        if (insertImageButton) {
            insertImageButton.addEventListener('click', onInsertImageClick);
        } else {
            console.warn('找不到 insert-license-button，图片插入按钮未绑定');
        }
    };

    window.Asc.plugin.button = function (id) { };

})(window, undefined);
