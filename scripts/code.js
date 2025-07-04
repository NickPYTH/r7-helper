(function (window, undefined) {
    var data = {
        rangeA: [],
        rangeB: [],
        range: []
    };
    window.Asc.plugin.init = function () {
        // Модуль Сравнение. Выбор диапазона А
        document.getElementById("compareSelectRangeABtn").onclick = function () {
            window.Asc.plugin.callCommand(function () {
                let result = {rangeA: [], rangeAText: ""}
                result.rangeAText = Api.GetSelection().Address;
                Api.GetSelection().ForEach(cell => {
                    //var oColor = Api.CreateColorFromRGB(255, 111, 61);
                    cell.FillColor = 'NoFill';
                    result.rangeA.push({
                        value: cell.GetValue(),
                        col: cell.GetCol(),
                        row: cell.GetRow()
                    })
                });
                //Api.GetSelection().SetValue("selected");
                return JSON.stringify(result);
            }, false, true, (range => {
                range = JSON.parse(range);
                data.rangeA = range.rangeA;
                document.getElementById("compareSelectRangeAText").textContent = "Д1: " + range.rangeAText.replaceAll("$", "");
            }));
        };
        // Модуль Сравнение. Выбор диапазона Б
        document.getElementById("compareSelectRangeBBtn").onclick = function () {
            window.Asc.plugin.callCommand(function () {
                let result = {rangeB: [], rangeBText: ""}
                result.rangeBText = Api.GetSelection().Address;
                Api.GetSelection().ForEach(cell => {
                    result.rangeB.push({
                        value: cell.GetValue(),
                        col: cell.GetCol(),
                        row: cell.GetRow()
                    })
                });
                return JSON.stringify(result);
            }, false, true, (range => {
                range = JSON.parse(range);
                data.rangeB = range.rangeB;
                document.getElementById("compareSelectRangeBText").textContent = "Д2: " + range.rangeBText.replaceAll("$", "");
            }));
        };
        // Модуль Сравнение. Выполнение сравнения
        document.getElementById("compareStartBtn").onclick = function () {
            localStorage.setItem('data', JSON.stringify(data))
            let compareMode = document.getElementById("compareSelectMode");
            switch (compareMode.value) {  // Выбор варианта сравнения
                case "1":  // Режим 1: Найти строки из Д1, совпадающие с Д2
                    window.Asc.plugin.callCommand(function () {
                        const convertNumberToLetter = (number) => {
                            switch (number) {
                                case 1:
                                    return 'A';
                                case 2:
                                    return 'B';
                                case 3:
                                    return 'C';
                            }
                            return '';
                        }
                        let data = JSON.parse(localStorage.getItem('data'));
                        data.rangeA.forEach(elA => {
                            data.rangeB.filter(elB => elA.value === elB.value).forEach(el => {
                                Api.GetRange(`${convertNumberToLetter(el.col)}${el.row}`).FillColor = Api.CreateColorFromRGB(255, 0, 0);
                            })
                        })
                    }, false, true, (result => {
                        localStorage.clear();
                    }));
                    break;
                case "2": // Режим 2: Найти строки из Д2, совпадающие с Д1
                    window.Asc.plugin.callCommand(function () {
                        const convertNumberToLetter = (number) => {
                            switch (number) {
                                case 1:
                                    return 'A';
                                case 2:
                                    return 'B';
                                case 3:
                                    return 'C';
                            }
                            return '';
                        }
                        let data = JSON.parse(localStorage.getItem('data'));
                        console.log(data)
                        data.rangeB.forEach(elB => {
                            data.rangeA.filter(elA => elB.value === elA.value).forEach(el => {
                                Api.GetRange(`${convertNumberToLetter(el.col)}${el.row}`).FillColor = Api.CreateColorFromRGB(255, 0, 0);
                            })
                        })
                    }, false, true, (result => {
                        localStorage.clear();
                    }));
                    break;
                case "3":  // Режим 3: Найти строки из Д1, которых нет в Д2
                    window.Asc.plugin.callCommand(function () {
                        const convertNumberToLetter = (number) => {
                            switch (number) {
                                case 1:
                                    return 'A';
                                case 2:
                                    return 'B';
                                case 3:
                                    return 'C';
                            }
                            return '';
                        }
                        let data = JSON.parse(localStorage.getItem('data'));
                        data.rangeA.forEach(elA => {
                            let searchedEl = data.rangeB.find(elB => elA.value === elB.value);
                            if (!searchedEl) {
                                Api.GetRange(`${convertNumberToLetter(elA.col)}${elA.row}`).FillColor = Api.CreateColorFromRGB(255, 0, 0);
                            }
                        })
                    }, false, true, (result => {
                        localStorage.clear();
                    }));
                    break;
                case "4":  // Режим 4: Найти строки из Д2, которых нет в Д1
                    window.Asc.plugin.callCommand(function () {
                        const convertNumberToLetter = (number) => {
                            switch (number) {
                                case 1:
                                    return 'A';
                                case 2:
                                    return 'B';
                                case 3:
                                    return 'C';
                            }
                            return '';
                        }
                        let data = JSON.parse(localStorage.getItem('data'));
                        data.rangeB.forEach(elB => {
                            let searchedEl = data.rangeA.find(elA => elB.value === elA.value);
                            if (!searchedEl) {
                                Api.GetRange(`${convertNumberToLetter(elB.col)}${elB.row}`).FillColor = Api.CreateColorFromRGB(255, 0, 0);
                            }
                        })
                    }, false, true, (result => {
                        localStorage.clear();
                    }));
                    break;
            }
        };
        // Модуль Сравнение. Очистка выделенного диапазона
        document.getElementById("compareClearRangeBtn").onclick = function () {
            window.Asc.plugin.callCommand(function () {
                Api.GetSelection().ForEach(cell => {
                    cell.FillColor = 'NoFill';
                });
            }, false, true);
        };
        // Модуль Замена формул в значения. Замена в выделенном диапазоне
        document.getElementById("castFormulasToValueSelectRangeBtn").onclick = function () {
            window.Asc.plugin.callCommand(function () {
                Api.GetSelection().ForEach(cell => {
                    cell.SetValue(cell.GetValue());
                });
            }, false, true);
        };
        // Модуль Дубликаты и уникальные значения. Выбор диапазона
        document.getElementById("duplicatesSelectRangeBtn").onclick = function () {
            window.Asc.plugin.callCommand(function () {
                let result = {range: [], rangeText: ""}
                result.rangeText = Api.GetSelection().Address;
                Api.GetSelection().ForEach(cell => {
                    result.range.push({
                        value: cell.GetValue(),
                        col: cell.GetCol(),
                        row: cell.GetRow()
                    })
                });
                return JSON.stringify(result);
            }, false, true, (range => {
                range = JSON.parse(range);
                data.range = range.range;
                document.getElementById("duplicatesRangeText").textContent = "Выбран диапазон: " + range.rangeText.replaceAll("$", "");
            }));
        };
        // Модуль Дубликаты и уникальные значения. Выполнение поиска
        document.getElementById("duplicatesStartBtn").onclick = function () {
            localStorage.setItem('data', JSON.stringify(data))
            let compareMode = document.getElementById("duplicatesSelectMode");
            switch (compareMode.value) {  // Выбор варианта сравнения
                case "1":  // Режим 1: Первые вхождения каждого элемента
                    window.Asc.plugin.callCommand(function () {
                        const convertNumberToLetter = (number) => {
                            switch (number) {
                                case 1:
                                    return 'A';
                                case 2:
                                    return 'B';
                                case 3:
                                    return 'C';
                            }
                            return '';
                        }
                        let data = JSON.parse(localStorage.getItem('data'));
                        let result = []
                        data.range.forEach((el, index) => {
                            if (result.find(e => e.value === el.value) === undefined) {
                                result.push(el)
                            }
                        })
                        result.forEach(el => {
                            Api.GetRange(`${convertNumberToLetter(el.col)}${el.row}`).FillColor = Api.CreateColorFromRGB(255, 0, 0);
                        })
                    }, false, true, (() => {
                        localStorage.clear();
                    }));
                    break;
                case "2":  // Режим 2: Дубликаты (все вхождения, кроме первого)
                    window.Asc.plugin.callCommand(function () {
                        const convertNumberToLetter = (number) => {
                            switch (number) {
                                case 1:
                                    return 'A';
                                case 2:
                                    return 'B';
                                case 3:
                                    return 'C';
                            }
                            return '';
                        }
                        let data = JSON.parse(localStorage.getItem('data'));
                        let uniq = []
                        let duplicates = []
                        data.range.forEach((el) => {  // Заполняем уникальные
                            if (uniq.find(e => e.element.value === el.value) === undefined) {
                                uniq.push({
                                    element: el,
                                    color: Api.CreateColorFromRGB(Math.floor(Math.random() * (255)), Math.floor(Math.random() * (255)), Math.floor(Math.random() * (255)))
                                })
                            }
                        })
                        data.range.forEach((el) => {  // Заполняем дубликаты, кроме 1 вхождения (1ые это уникальные)
                            if (uniq.find(ue => ue.element.value === el.value && el.row === ue.element.row && el.col === ue.element.col) === undefined) {
                                let tmp = uniq.find(ue => ue.element.value === el.value);
                                if (tmp)
                                    duplicates.push({element: el, color: tmp.color})
                            }
                        })
                        duplicates.forEach(el => {  // Красим дубли
                            Api.GetRange(`${convertNumberToLetter(el.element.col)}${el.element.row}`).FillColor = el.color;
                        })
                    }, false, true, (() => {
                        localStorage.clear();
                    }));
                    break;
                case "3":  // Режим 3: Уникальные
                    window.Asc.plugin.callCommand(function () {
                        const convertNumberToLetter = (number) => {
                            switch (number) {
                                case 1:
                                    return 'A';
                                case 2:
                                    return 'B';
                                case 3:
                                    return 'C';
                            }
                            return '';
                        }
                        let data = JSON.parse(localStorage.getItem('data'));
                        let uniq = []
                        data.range.forEach((el) => {  // Заполняем уникальные
                            if (data.range.filter(e => e.value === el.value).length === 1) {
                                uniq.push({
                                    element: el,
                                    color: Api.CreateColorFromRGB(Math.floor(Math.random() * (255)), Math.floor(Math.random() * (255)), Math.floor(Math.random() * (255)))
                                })
                            }
                        })
                        uniq.forEach(el => {  // Красим уникальные
                            Api.GetRange(`${convertNumberToLetter(el.element.col)}${el.element.row}`).FillColor = el.color;
                        })
                    }, false, true, (() => {
                        localStorage.clear();
                    }));
                    break;
                case "4":  // Режим 4: Повторяющиеся (встречающиеся больше одного раза)
                    window.Asc.plugin.callCommand(function () {
                        const convertNumberToLetter = (number) => {
                            switch (number) {
                                case 1:
                                    return 'A';
                                case 2:
                                    return 'B';
                                case 3:
                                    return 'C';
                            }
                            return '';
                        }
                        let data = JSON.parse(localStorage.getItem('data'));
                        let repeats = []
                        data.range.forEach((el) => {  // Заполняем уникальные
                            if (data.range.filter(e => e.value === el.value).length > 1) {
                                let tmp = repeats.find(r => r.element.value === el.value)
                                if (tmp)
                                    repeats.push({
                                        element: el,
                                        color: tmp.color
                                    })
                                else
                                    repeats.push({
                                        element: el,
                                        color: Api.CreateColorFromRGB(Math.floor(Math.random() * (255)), Math.floor(Math.random() * (255)), Math.floor(Math.random() * (255)))
                                    })
                            }
                        })
                        repeats.forEach(el => {  // Красим уникальные
                            Api.GetRange(`${convertNumberToLetter(el.element.col)}${el.element.row}`).FillColor = el.color;
                        })
                    }, false, true, (() => {
                        localStorage.clear();
                    }));
                    break;
            }
        };
        // Модуль Модуль Дубликаты и уникальные значения. Очистка выделенного диапазона
        document.getElementById("duplicatesClearRangeBtn").onclick = function () {
            window.Asc.plugin.callCommand(function () {
                Api.GetSelection().ForEach(cell => {
                    cell.FillColor = 'NoFill';
                });
            }, false, true);
        };
    };

    window.Asc.plugin.button = function () {
        this.executeCommand("close", "");
    };

})(window, undefined);