// 折线图表一
var MONTHS = ['0:00', '2:00', '4:00', '6:00', '8:00', '10:00', '12:00', '14:00', '16:00', '18:00', '20:00', '22:00'];
var config1 = {
    type: 'line',
    data: {
        labels: ['0:00', '2:00', '4:00', '6:00', '8:00', '10:00', '12:00', '14:00', '16:00', '18:00', '20:00', '22:00'],
        datasets: [
            {
                label: 'CPU使用率',
                backgroundColor: window.chartColors.red,
                borderColor: window.chartColors.red,
                data: [
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor()
                ],
                fill: false,
            },
            {
                label: '服务器带宽',
                fill: false,
                backgroundColor: window.chartColors.blue,
                borderColor: window.chartColors.blue,
                data: [
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor()
                ],
            }, {
                label: '服务器网络',
                fill: false,
                backgroundColor: window.chartColors.orange,
                borderColor: window.chartColors.orange,
                data: [
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor()
                ],
            },
            {
                label: '网站访问量',
                fill: false,
                backgroundColor: window.chartColors.green,
                borderColor: window.chartColors.green,
                data: [
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor(),
                    randomScalingFactor()
                ],
            }
        ]
    },
    options: {
        responsive: true,
        title: {
            display: true,
            text: '服务器使用状态'
        },
        tooltips: {
            mode: 'index',
            intersect: false,
        },
        hover: {
            mode: 'nearest',
            intersect: true
        },
        scales: {
            x: {
                display: true,
                scaleLabel: {
                    display: true,
                    labelString: 'Month'
                }
            },
            y: {
                display: true,
                scaleLabel: {
                    display: true,
                    labelString: 'Value'
                }
            }
        }
    }
};


document.getElementById('randomizeData').addEventListener('click', function () {
    config1.data.datasets.forEach(function (dataset) {
        dataset.data = dataset.data.map(function () {
            return randomScalingFactor();
        });

    });

    window.myLine.update();
});

var colorNames = Object.keys(window.chartColors);
document.getElementById('addDataset').addEventListener('click', function () {
    var colorName = colorNames[config1.data.datasets.length % colorNames.length];
    var newColor = window.chartColors[colorName];
    var newDataset = {
        label: 'Dataset ' + config1.data.datasets.length,
        backgroundColor: newColor,
        borderColor: newColor,
        data: [],
        fill: false
    };

    for (var index = 0; index < config1.data.labels.length; ++index) {
        newDataset.data.push(randomScalingFactor());
    }

    config1.data.datasets.push(newDataset);
    window.myLine.update();
});

document.getElementById('addData').addEventListener('click', function () {
    if (config1.data.datasets.length > 0) {
        var month = MONTHS[config1.data.labels.length % MONTHS.length];
        config1.data.labels.push(month);

        config1.data.datasets.forEach(function (dataset) {
            dataset.data.push(randomScalingFactor());
        });

        window.myLine.update();
    }
});

document.getElementById('removeDataset').addEventListener('click', function () {
    config1.data.datasets.splice(0, 1);
    window.myLine.update();
});

document.getElementById('removeData').addEventListener('click', function () {
    config1.data.labels.splice(-1, 1); // remove the label first

    config1.data.datasets.forEach(function (dataset) {
        dataset.data.pop();
    });

    window.myLine.update();
});



//  分析图

var randomScalingFactor = function () {
    return Math.round(Math.random() * 100);
};

var chartColors = window.chartColors;
var color = Chart.helpers.color;
var config2 = {
    type: 'polarArea',
    data: {
        datasets: [{
            data: [
                randomScalingFactor(),
                randomScalingFactor(),
                randomScalingFactor(),
                randomScalingFactor(),
                randomScalingFactor(),
            ],
            backgroundColor: [
                color(chartColors.red).alpha(0.5).rgbString(),
                color(chartColors.orange).alpha(0.5).rgbString(),
                color(chartColors.yellow).alpha(0.5).rgbString(),
                color(chartColors.green).alpha(0.5).rgbString(),
                color(chartColors.blue).alpha(0.5).rgbString(),
            ],
            label: 'My dataset' // for legend
        }],
        labels: [
            '广告收入',
            '用户打赏',
            '商品出售',
            '流量导向',
            '网站抽成'
        ]
    },
    options: {
        responsive: true,
        legend: {
            position: 'right',
        },
        title: {
            display: true,
            text: '网站收入占比情况图'
        },
        scale: {
            ticks: {
                beginAtZero: true
            },
            reverse: false
        },
        animation: {
            animateRotate: false,
            animateScale: true
        }
    }
};


document.getElementById('randomizeData2').addEventListener('click', function () {
    config2.data.datasets.forEach(function (piece, i) {
        piece.data.forEach(function (value, j) {
            config2.data.datasets[i].data[j] = randomScalingFactor();
        });
    });
    window.myPolarArea.update();
});

var colorNames = Object.keys(window.chartColors);
document.getElementById('addData2').addEventListener('click', function () {
    if (config2.data.datasets.length > 0) {
        config2.data.labels.push('data #' + config2.data.labels.length);
        config2.data.datasets.forEach(function (dataset) {
            var colorName = colorNames[config2.data.labels.length % colorNames.length];
            dataset.backgroundColor.push(window.chartColors[colorName]);
            dataset.data.push(randomScalingFactor());
        });
        window.myPolarArea.update();
    }
});
document.getElementById('removeData2').addEventListener('click', function () {
    config2.data.labels.pop(); // remove the label first
    config2.data.datasets.forEach(function (dataset) {
        dataset.backgroundColor.pop();
        dataset.data.pop();
    });
    window.myPolarArea.update();
});

// 服务器

window.onload = function () {
    var ctx1 = document.getElementById('fircanvas').getContext('2d');
    window.myLine = new Chart(ctx1, config1);
    // 分析图
    var ctx2 = document.getElementById('seccanvas');
    window.myPolarArea = new Chart(ctx2, config2);
};

// 时间插入

function getTime() {
    let systemtime = document.getElementsByClassName('systemtime');
    let time = new Date()
    let year = time.getFullYear();
    let month = time.getMonth();
    let day = time.getDay();
    let hour = time.getHours();
    let minutes = time.getMinutes();
    let seconds = time.getSeconds();
    for (let i = 0; i < systemtime.length; i++) {
        systemtime[i].innerText = year + "年" + month + "月" + day + "日" + hour + "时" + minutes + "分" + seconds + "秒";
    }
}

function refresh() {
    setInterval(getTime, 1000);
}

refresh();




