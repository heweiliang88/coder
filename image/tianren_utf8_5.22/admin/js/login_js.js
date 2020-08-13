// JavaScript Document
var bg_img = $('lay_bg_img'),
bgArr = [];
for (var i = 0,
len = p_bgPic.length; i < len; i++) {
  if (p_bgPic[i].length > 0) {
    bgArr.push(p_bgPic[i]);
  }
}
var bg_Data = bgArr[Math.floor(Math.random() * (bgArr.length))],
bg_type = bg_Data[0],
ft_col = bg_Data[2] || 0;

window.QZFL = window.QZONE = window.QZFL || window.QZONE || {};
QZFL.dom = {
  getById: function(id) {
    return document.getElementById(id);
  },
  getClientHeight: function(doc) {
    var _doc = doc || document;
    return _doc.compatMode == "CSS1Compat" ? _doc.documentElement.clientHeight: _doc.body.clientHeight;
  },
  getClientWidth: function(doc) {
    var _doc = doc || document;
    return _doc.compatMode == "CSS1Compat" ? _doc.documentElement.clientWidth: _doc.body.clientWidth;
  }
};
QZFL.css = {
  _reClassToken: /\s+/,
  addClassName: function(elem, names) {
    var _s = QZFL.css;
    return names && ((elem && elem.classList && !_s._reClassToken.test(names)) ? elem.classList.add(names) : _s.updateClassName(elem, null, names));
  },
  removeClassName: function(elem, names) {
    var _s = QZFL.css;
    return names && ((elem && elem.classList && !_s._reClassToken.test(names)) ? elem.classList.remove(names) : _s.updateClassName(elem, names));
  },
  updateClassName: function(elem, removeNames, addNames) {
    if (!elem || elem.nodeType != 1) {
      return "";
    }
    var oriName = elem.className,
    _s = QZFL.css,
    ar, b;
    if (removeNames && typeof(removeNames) == 'string' || addNames && typeof(addNames) == 'string') {
      if (removeNames == '*') {
        oriName = '';
      } else {
        ar = oriName.split(_s._reClassToken);
        var i = 0,
        l = ar.length,
        n;
        oriName = {};
        for (; i < l; ++i) {
          ar[i] && (oriName[ar[i]] = true);
        }
        if (addNames) {
          ar = addNames.split(_s._reClassToken);
          l = ar.length;
          for (i = 0; i < l; ++i) { (n = ar[i]) && !oriName[n] && (b = oriName[n] = true);
          }
        }
        if (removeNames) {
          ar = removeNames.split(_s._reClassToken);
          l = ar.length;
          for (i = 0; i < l; i++) { (n = ar[i]) && oriName[n] && (b = true) && delete oriName[n];
          }
        }
      }
      if (b) {
        ar.length = 0;
        for (var k in oriName) {
          ar.push(k);
        }
        oriName = ar.join(' ');
        elem.className = oriName;
      }
    }
    return oriName;
  }
};

if (!window.$) {
  var $ = QZFL.dom.getById;
}
QZONE.LoginPage = {
  bootStrap: function() {
    var lp = QZONE.LoginPage,
    sl_url = $('small_url');
    if (bg_type == 0) {
      QZFL.css.addClassName(document.body, 'mode_dark');
    }
    bg_img.onload = function() {
      QZFL.css.addClassName(bg_img, 'lay_background_img_fade_out');
      lp.resizeBackground();
    }
    window.onload = function() {
      lp.resizeBackground();
      lp.setLoginDivTop();
    };

    //		TCISD.pv('ihome.qzone.qq.com','login/i');
  },
  resizeBackground: function() {
    var bg = $('lay_bg'),
    bg_img = $('lay_bg_img'),
    cw = QZFL.dom.getClientWidth(),
    ch = QZFL.dom.getClientHeight(),
    iw = bg_img.width,
    ih = bg_img.height;

    bg.style.width = cw + "px";
    bg.style.height = ch + "px";

    if (cw / ch > iw / ih) {
      var new_h = cw * ih / iw,
      imgTop = (ch - new_h) / 2;
      bg_img.style.width = cw + "px";
      bg_img.style.height = new_h + "px";
      bg_img.style.top = imgTop + "px";
      bg_img.style.left = "";
    } else {
      var new_w = ch * iw / ih,
      imgLeft = (cw - new_w) / 2;
      bg_img.style.width = new_w + "px";
      bg_img.style.height = ch + "px";
      bg_img.style.left = imgLeft + "px";
      bg_img.style.top = "";
    }
  },
  setLoginDivTop: function() {
    var dom_height = QZFL.dom.getClientHeight();

    if (window.ActiveXObject && (navigator.userAgent.indexOf('MSIE 6.0') > -1) && dom_height < 600) {
      $('lay').style.height = '600px';
    } else {
      $('lay').style.height = '';
    }

    if (dom_height < 820) {
      var y1 = 820 - dom_height,
      change_top = 100 - y1,
      change_top = change_top > 0 ? change_top: 0;
      //			$('login_div').style.top = change_top + "px";
    } else {
      //			$('login_div').style.top = "100px";
    }
  }
};
QZONE.LoginPage.bootStrap();

function mrc() {
var strs = new Array('只要还有明天，今天就永远是起跑线。','火把倒下，火焰依然向上。','成功就是你被击落到失望的深渊之后反弹得有多高。','雨是云的梦，云是雨的前世。','树没有眼睛，落叶却是飘落的眼泪。','天将降大任于斯人也，天不降大任，你不还是斯人吗？','别刚从底层混起，就混得特别没底。','有时候对一个作家而言，真正的奖赏不是诺贝尔奖，而是盗版。','腾不出时间娱乐，早晚会被迫腾出时间生病。','大家都不急，我急什么？大家都急了，我能不急？这就是工作与奖金的区别。','你不能因为自己是刘翔，就看不起哪些参加全民健身的。','每个人最终和自己越长越像。','思念是没有翅膀的鸟，在心里盲目的停留不前。','阳光落在春的枝头，日子便绿了。','旁若无人走自己的路，不是撞到别人，就是被别人撞到。','只要比竞争对手活得长，你就赢了。','最好的，不一定是最合适的；最合适的，才是真正最好的。','历史，只有人名真的；小说，只有人名是假的。','你不是诗人，但可以诗意地生活。如果能够诗意地生活，那你就是诗人。','眼睛为你下着雨，心却为你撑着伞。  ','爱或被爱，不如相爱。','美丽是危险的，象因牙死，狐因皮亡。','一个人害怕的事，往往是他应该做的事。','我们很早在乎自己是否有机会发言，却不在乎是否有人愿意听。','把影子，腌起来，风干，老的时候，下酒。','青椒肉丝和肉丝青椒是同一道菜，关键是谁来炒。','除非你是三栖动物，否则，总有一个空间不属于你。','与其给鱼一双翅膀，不如给鱼一方池塘。','人生的上半场打不好没关系，还有下半场，只要努力。','为了照亮夜空，星星才站在天空的高处。','缺点也是点，点到为止。','如果你坚信石头会开花，那么开花的不仅仅是石头。','没有月色的夜里，萤火虫用心点亮漫天的星。','诚信总会给你带来成功，但可能是下一站。','你没有摘到的，只是春天里的一朵花，整个春天还是你的。','据我测算，还可以退五十步，但生活只有五步。','最深沉的感情，往往以最冷漠的方式表达出来。','脚是大地上飞翔的翅膀。','苦难是化了装的幸福。','当一个人有了想飞的梦想，哪怕爬着，也没有不站起来的理由。','当所有的爱熄灭，还可以点燃自己，让心亮着。','泪水不是为了排除外在的悲伤，而是为了自由的哭泣。','当你一门心思撞南墙的时候，南墙就消失了。','成功其实很简单，就是当你坚持不住的时候，再坚持一下。','爱慕月光的人并不将它据为己有，能远远望着，就已足够。','想念，滴在左手凝固成寂寞，落在右手化为牵挂。','长得漂亮是优势，活得漂亮是本事。','死，可以明志；生，却可践志。','我虽然不同意你的观点，但我誓死捍卫你说话的权力。','人生没有彩排，每天都是现场直播！','人生来如风雨，去如微尘。','每个人都想和别人不一样，结果是每个人都一样。','有时候，不是对方不在乎你，而是你把对方看的太重。','忍无可忍，就重新再忍。','美丽让男人停下，智慧让男人留下。','天使之所以会飞，是因为她们把自己看得很轻。','没有人值得你流泪，值得让你这么做的人不会让你哭泣。','对于世界，你可能只是一个人，但对于某个人，你却是整个世界。','每个人出生的时候都是原创，可悲的是很多人渐渐都成了盗版。','世界上只有想不通的人，没有走不通的路。','难过时，吃一粒糖，告诉自己生活是甜的！','养成每天写点什么的习惯，哪怕是记录，哪怕只言片语。','成功便是站起比倒下多一次。','状态是干出来的，而不是等出来的。','使人疲惫的不是远方的高山，而是鞋里的一粒沙子。','给猴一棵树，给虎一座山。','要随波逐浪，不可随波逐流。','前方无绝路，希望在转角。','世上没有绝望的处境，只有对处境绝望的人。','心有阳光，抬望远方，明媚敞亮，静下来，找个时间与自己对话。','谤言只是挖在你身后的陷阱，只要你一直勇往直前，陷阱再多也不一定伤害到你。','工作是工作，生活是生活。工作上，自己扮演的是职务角色；生活里，自己扮演的是自我。工作只是生命中一小部分，切勿把谋生工具当成人生的全部生活！','上苍不会给你快乐也不会给你痛苦，它只会给你真实的生活。有人忍受不了生活的平淡而死去，却不知道生命本身就是奇迹！','生活，就象一个无形的天平，站在上面的每个人都有可能走极端，但这最终都是为了寻找一个平衡的支点，使自己站的更稳，走的更好，活的更精彩!','穷不一定思变，应该是思富思变。','如果心胸不似海，又怎能有海一样的事业。','人之所以能，是相信能。','不要等着别人发现你，你应该先把自己推出去！','成功与不成功之间有时距离很短——只要后者再向前几步。','对于最有能力的领航人风浪总是格外的汹涌。','挫折其实就是迈向成功所应缴的学费。','如同磁铁吸引四周的铁粉，热情也能吸引周围的人，改变周围的情况。','你可以很有个性，但某些时候请收敛。','好的想法是十分钱一打，真正无价的是能够实现这些想法的人。','欲望以提升热忱，毅力以磨平高山。','最惬意的时候，往往是失败的开始。','行动不一定带来快乐，而无行动则决无快乐。','要想出头，必须先学会出丑。','未遭拒绝的成功决不会长久。','寒冷到了极致时，太阳就要光临。','只有千锤百炼，才能成为好钢。','环境永远不会十全十美，消极的人受环境控制，积极的人却控制环境。','外在压力增加时，就应增强内在的动力。','你不勇敢，没人替你坚强。','每一发奋努力的背后，必有加倍的赏赐。','投资知识是明智的，投资网络中的知识就更加明智。','做对的事情比把事情做对重要。','最忙的时候，学的东西可能最多。','成功呈概率分布，关键是你能不能坚持到成功开始呈现的那一刻。','如果寒暄只是打个招呼就了事的话，那与猴子的呼叫声有什么不同呢？事实上，正确的寒暄必须在短短一句话中明显地表露出你他的关怀。','如果你孩子贫穷是你父母的罪恶，因为在他小时候，你没有给他正确的人生观。','自己打败自己的远远多于比别人打败的。','这个世界并不是掌握在那些嘲笑者的手中，而恰恰掌握在能够经受得住嘲笑与批评并不断往前走的人手中。','一个讯息从地球这一端到另一端只需一秒，而一个观念从脑外传到脑里却需要一年，三年甚至十五年。','行动是治愈恐惧的良药，而犹豫拖延将不断滋养恐惧。','空想会想出很多绝妙的主意，但却办不成任何事情。','对的，坚持；错的，放弃！','生命对某些人来说是美丽的，这些人的一生都为某个目标而奋斗。推销产品要针对顾客的心，不要针对顾客的头。','苦想没盼头，苦干有奔头。','积极者相信只有推动自己才能推动世界，只要推动自己就能推动世界。','只要我们能梦想的，我们就能实现。','忍耐力较诸脑力，尤胜一筹。','大多数人想要改造这个世界，但却罕有人想改造自己。','与其做一个有价钱的人，不如做一个有价值的人；与其做一个忙碌的人，不如做一个有效率的人。','没有人富有得可以不要别人的帮助，也没有人穷得不能在某方面给他人帮助。','自己打败自己是最可悲的失败，自己战胜自己是最可贵的胜利。','前有阻碍，奋力把它冲开，运用炙热的激情，转动心中的期待，血在澎湃，吃苦流汗算什么。','成功的信念在人脑中的作用就如闹钟，会在你需要时将你唤醒。','个人魅力：和谐快乐坚毅品质心境仪表谈吐，内涵修为。','如果你希望成功，以恒心为良友，以经验为参谋，以小心为兄弟，以希望为哨兵。','人生伟业的建立，不在能知，乃在能行。','没有不老的誓言，没有不变的承诺，踏上旅途，义无反顾！','我们可以输在人生起跑点，但绝不能输在人生转折点。','投资知识是明智的，投资网络中的知识就更加明智。没有天生的信心，只有不断培养的信心。','征服畏惧建立自信的最快最确实的方法，就是去做你害怕的事，直到你获得成功的经验。','吃别人所不能吃的苦，忍别人不能忍的气，做别人所不能做的事，就能享受别人所不能享受的一切。','没有一种不通过蔑视忍受和奋斗就可以征服的命运。','生命的辉煌，拒绝的不是平凡，而是平庸！所以，春风得意时多些缅想，只要别背叛美丽的的初衷；窘迫失意时多些憧憬，只要别虚构不醒的苦梦！','生命之灯因热情而点燃，生命之舟因拼搏而前行。','自古成功在尝试。','人性最可怜的就是：我们总是梦想着天边的一座奇妙的玫瑰园，而不去欣赏今天就开在我们窗口的玫瑰。','少说多做，句句都会得到别人的重视；多说少做，句句都会受到别人的忽视。','智者一切求自己，愚者一切求他人。','人若把自己框在一定的范围内，就容易限制了自己的思维和格局。','凡事要三思，但比三思更重要的是三思而行。','好咖啡要和朋友一起品尝，好机会也要和朋友一起分享。','如果我们做与不做都会有人笑，如果做不好与做得好还会有人笑，那么我们索性就做得更好，来给人笑吧！','每天早上醒来，你荷包里的最大资产是二十四个小时——你生命宇宙中尚未制造的材料。','如果要挖井，就要挖到水出为止。','为明天做准备的最好方法就是集中你所有智慧，所有的热忱，把今天的工作做得尽善尽美，这就是你能应付未来的唯一方法。','只要你把头抬高点，别人就不会看低你，记住，人们是看不起低着头的人的。','一个能从别人的观念来看事情，能了解别人心灵活动的人，永远不必为自己的前途担心。','靠山山会倒，靠水水会流，靠自己永远不倒。','世上最重要的事，不在于我们在何处，而在于我们朝着什么方向走。'); 
var strs2 = Math.floor(Math.random() * strs.length); 
document.getElementById('hearten').innerHTML = strs[strs2];
}
window.onload = mrc;
