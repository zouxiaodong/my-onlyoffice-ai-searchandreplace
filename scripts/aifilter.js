(function(window, undefined) {
    var Api;

    var pluginConfig = {
        backendUrl: 'http://192.168.4.43:8001/api/fill-document'
    };

    function updateStatus(message, isError = false) {
        var statusElement = document.getElementById('status-message');
        if (statusElement) {
            statusElement.textContent = message;
            statusElement.classList.toggle('error', isError);
        }
    }

    function onButtonClick() {
        var button = document.getElementById('fill-button');
        if (button) button.disabled = true;
        updateStatus('正在读取文档内容...');

        // 第一次callCommand：获取文档内容并返回
        window.Asc.plugin.callCommand(function() {
            var oDocument = Api.GetDocument();
            if (!oDocument) return '';

            var text = oDocument.GetText() || '';
            console.log('callCommand中获取文档内容:', text);

            // 返回文档文本给回调函数
            return text;
        }, false, false, function(returnedText) {
            // 这里的returnedText就是callCommand中返回的文档文本
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
                headers: {'Content-Type': 'application/json'},
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
                Asc.scope.replacements = JSON.stringify(replacements); // 存储为字符串

                window.Asc.plugin.callCommand(function() {
                    var reps = JSON.parse(Asc.scope.replacements); // 解析为对象
                    var oDoc = Api.GetDocument();
                    if (!oDoc) return;

                    // 遍历替换内容，执行替换操作
                    reps.forEach(function(replacement) {
                        var target = replacement.target;
                        var value = replacement.value;
                        if (target && value !== undefined && value !== null) {
                            console.log('替换:', target, '->', value);

                            // 使用 SearchAndReplace 执行替换
                            oDoc.SearchAndReplace({
                                searchString: target,
                                replaceString: value
                            });
                        }
                    });

                }, true, true, function() {
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

    window.Asc.plugin.init = function(api) {
        Api = api;
        var button = document.getElementById('fill-button');
        if (button) {
            button.addEventListener('click', onButtonClick);
        } else {
            console.error('找不到 fill-button');
        }
    };

    window.Asc.plugin.button = function(id) {};

})(window, undefined);