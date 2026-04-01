try {
    var fso = new ActiveXObject("Scripting.FileSystemObject");

    // Setup file paths
    var dirStr = unescape(location.href).split('/').slice(0, -1).join('/');
    var currentDir = dirStr.replace("file:///", "").replace(/\//g, "\\");

    var AppFiles = {
        config: currentDir + "\\config.json",
        history: currentDir + "\\history.json",
        questions: currentDir + "\\questions.json"
    };

    var $id = function (id) { return document.getElementById(id); };

    function readJsonFile(filePath, defaultJson) {
        if (!fso.FileExists(filePath)) return JSON.parse(defaultJson);
        var stream = new ActiveXObject("ADODB.Stream");
        stream.Type = 2; stream.Charset = "utf-8";
        try {
            stream.Open();
            stream.LoadFromFile(filePath);
            var text = stream.ReadText();
            stream.Close();
            return JSON.parse(text);
        } catch (e) {
            return JSON.parse(defaultJson);
        }
    }

    function writeJsonFile(filePath, obj) {
        var str = new ActiveXObject("ADODB.Stream");
        str.Type = 2; str.Charset = "utf-8"; str.Open();
        str.WriteText(JSON.stringify(obj, null, 2));
        str.SaveToFile(filePath, 2);
        str.Close();
    }

    var config = readJsonFile(AppFiles.config, '{"algorithm": "weakness", "theme": "dark"}');
    var questions = readJsonFile(AppFiles.questions, '[]');
    var hist = readJsonFile(AppFiles.history, '{"streak": 0, "totalCorrect": 0, "totalAnswered": 0, "records": {}, "lastSequentialId": 0}');
    var currentQuestion = null;

    function init() {
        // HTAのcaption="no"による最大化バグ対策（手動で全画面化）
        window.moveTo(0, 0);
        window.resizeTo(screen.width, screen.height);

        if (!questions.length) {
            $id('question-text').innerText = "エラー: 問題ファイル (questions.json) が見つからないか、空です。";
            return;
        }

        document.body.className = config.theme || "dark";

        var acc = hist.totalAnswered > 0 ? ((hist.totalCorrect / hist.totalAnswered) * 100).toFixed(1) : "0.0";
        $id('streak-stat').innerText = "連勝記録: 🔥 " + (hist.streak || 0);
        $id('accuracy-stat').innerText = "累計正答率: 🎯 " + acc + "%";

        currentQuestion = selectQuestion(config.algorithm);
        $id('question-text').innerText = currentQuestion.question;

        var optionsArea = $id('options-area');
        optionsArea.className = "options-grid";
        optionsArea.innerHTML = '';

        for (var i = 0; i < currentQuestion.options.length; i++) {
            var btn = document.createElement('button');
            btn.innerText = currentQuestion.options[i];
            btn.onclick = (function (idx) { return function () { handleAnswer(idx); }; })(i);
            optionsArea.appendChild(btn);
        }
    }

    function selectQuestion(algo) {
        if (algo === "sequential") {
            var nextIdx = 0, lastId = hist.lastSequentialId || 0;
            for (var i = 0; i < questions.length; i++) {
                if (questions[i].id > lastId) {
                    nextIdx = i; break;
                }
            }
            return questions[nextIdx];
        }

        if (algo === "weakness") {
            var weights = [], totalWeight = 0;
            for (var i = 0; i < questions.length; i++) {
                var rec = hist.records[questions[i].id.toString()] || { correct: 0, wrong: 0 };
                var weight = Math.max(0.2, 1.0 + (rec.wrong * 2.0) - (rec.correct * 0.5));
                weights.push(weight);
                totalWeight += weight;
            }
            var rand = Math.random() * totalWeight, cum = 0;
            for (var i = 0; i < questions.length; i++) {
                cum += weights[i];
                if (rand <= cum) return questions[i];
            }
        }

        return questions[Math.floor(Math.random() * questions.length)];
    }

    function handleAnswer(selectedIdx) {
        var isCorrect = (selectedIdx === currentQuestion.answerIndex);
        var qId = currentQuestion.id.toString();

        var buttons = $id('options-area').getElementsByTagName('button');
        for (var i = 0; i < buttons.length; i++) {
            buttons[i].disabled = true;
            if (i === currentQuestion.answerIndex) buttons[i].className = "correct";
            else if (i === selectedIdx && !isCorrect) buttons[i].className = "wrong";
        }

        var resMsg = $id('result-msg');
        resMsg.innerText = isCorrect ? "⭕ 正解！" : "❌ 不正解...";
        resMsg.className = isCorrect ? "msg-correct" : "msg-wrong";

        $id('explanation').innerText = currentQuestion.explanation || "";
        $id('result-area').style.display = "block";
        $id('close-area').style.display = "block";

        if (!hist.records[qId]) hist.records[qId] = { correct: 0, wrong: 0 };

        hist.totalAnswered = (hist.totalAnswered || 0) + 1;
        if (isCorrect) {
            hist.streak = (hist.streak || 0) + 1;
            hist.totalCorrect = (hist.totalCorrect || 0) + 1;
            hist.records[qId].correct++;
        } else {
            hist.streak = 0;
            hist.records[qId].wrong++;
        }
        hist.lastSequentialId = currentQuestion.id;

        try { writeJsonFile(AppFiles.history, hist); } catch (e) { console.log("Failed to save history:", e); }
    }

    document.onkeydown = function (e) { if ((e || window.event).keyCode === 27) return false; };
    window.onload = function () { setTimeout(init, 50); };

} catch (e) {
    document.write("<h1 style='color:red'>初期化エラー</h1><p>" + e.message + "</p>");
}
