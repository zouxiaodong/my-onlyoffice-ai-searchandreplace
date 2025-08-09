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

        // Ensure the image URL is valid
        if (!pluginConfig.licenseImageUrl) {
            console.log('图片URL无效');
            if (button) button.disabled = false;
            return;
        }

        // Pass the image URL to the sandbox variable
        Asc.scope.licenseImageUrl = pluginConfig.licenseImageUrl;

        window.Asc.plugin.callCommand(function () {
            var oDocument = Api.GetDocument();
            if (!oDocument) return;

            // Get the current paragraph where the cursor is located
            var oParagraph = oDocument.GetCurrentParagraph();
            if (!oParagraph) {
                console.log('无法获取光标位置');
                if (button) button.disabled = false;
                return;
            }
            // 1 mm = 36000 EMUs, 1 inch = 914400 EMUs.
            var width = 10 * 36000 * 10;  // 10 cm
            var height = 8 * 36000 * 10;;  // 8 cm
            // Add a picture content control at the cursor location
            var oPictureControl = oDocument.AddPictureContentControl(width, height);
            if (oPictureControl) {
                console.log('添加图片:', Asc.scope.licenseImageUrl);
                // Set the image for the picture content control using the provided URL
                oPictureControl.SetPicture(Asc.scope.licenseImageUrl);
                console.log('营业执照图片插入成功');
            } else {
                console.log('插入图片失败');
            }

        }, false, false, function () {
            console.log('图片插入过程完成');
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
