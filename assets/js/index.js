// 鼠标移动到nav显示,移出nav2秒后隐藏 ，如果在子导航中，则要经过移出才隐藏
$(function () {
    $('#dropdown').hover(function () {
        $('#dropdown-menu').stop().slideToggle()
    })
})

$(function () {
    $(window).scroll(function () {
        var topHeight = $(document).scrollTop();
        if (topHeight > 300) {
            $('.conColor').css('backgroundColor', 'rgba(55, 64, 85, 0.9)')
        }
        else {
            $('.conColor').css('backgroundColor', 'transparent')
        }
    })
})

// banner的拒绝996的功能 section存储 其他的不可选 localStorage 属性 只能选择一个
$(function () {
    let anum = 130348;
    let znum = 6292;
    let rnum = 3452;


    $('#statnumber > button:nth(0)> b').click(function () {
        anum++;
        $(this).html(anum)
    })
    $('#statnumber > button:nth(1) > b').click(function () {
        znum++;
        $(this).html(znum)
    })
    $('#statnumber > button:nth(2) > b').click(function () {
        rnum++;
        $(this).html(rnum)
    })
})


$(function () {
    $('body').running();
})


new WOW().init();

