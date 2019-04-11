Template.boardMenuPopup.events({
  'click .pop-over-list-a': function () {
    var blob;
    var filename;
    function c() {
      window.navigator.msSaveOrOpenBlob(blob, filename);
    }
    //导出成csv
    function JSONToCSV(JSONData, ReportName, ShowLabel) {

      var arrData = typeof JSONData != 'object' ? JSON.parse(JSONData) : JSONData;
      var CSV = '';
      if (ShowLabel) {
        var row = "";

        for (var index in arrData[0]) {
          row += index + ',';
        }
        row = row.slice(0, -1);
        CSV += row + '\r\n';
      }

      for (var i = 0; i < arrData.length; i++) {
        var row = "";
        for (var index in arrData[i]) {
          row += '"' + arrData[i][index] + '",';
        }
        row.slice(0, row.length - 1);
        //add a line break after each row
        CSV += row + '\r\n';
        // console.log(CSV);
      }


      if (CSV == '') {
        alert("数据有错误");
        return;
      }
      // var uri = 'data:text/csv;charset=utf-8,\uFEFF' + encodeURI(CSV);
      var link = document.createElement("a");
      link.id = "lnkDwnldLnk";
      link.name = "lnkDwnldLnk";
      link.setAttribute("style", "display:none");


      var ifdo = document.createElement("iframe");
      ifdo.id = "ifdiframe";
      ifdo.name = "ifdiframe";
      ifdo.setAttribute("style", "display:none");

      document.body.appendChild(link);
      document.body.appendChild(ifdo);

      var csv = CSV;

      try {
        blob = new Blob([csv], { type: 'text/csv;charset=utf-8,\uFEFF' });
      }
      catch (e) {
        blob = [csv];
      }

      var csvUrl = "";
      filename = ReportName;
      if (window.navigator.userAgent.indexOf("Chrome") >= 1 || window.navigator.userAgent.indexOf("Safari") >= 1) {
        csvUrl = window.webkitURL.createObjectURL(blob);
        // csvUrl = 'data:text/csv;charset=utf-8,\uFEFF' +csvUrl;
        link.setAttribute("download", filename);
        var uri = 'data:text/csv;charset=utf-8,\uFEFF' + encodeURI(csv);
        // link.setAttribute("href", csvUrl);
        link.setAttribute("href", uri);
        link.click();
      }
      if (window.navigator.userAgent.indexOf("Firefox") >= 1) {
        csvUrl = window.URL.createObjectURL(blob);
        link.setAttribute("download", filename);
        var uri = 'data:text/csv;charset=utf-8,\uFEFF' + encodeURI(csv);
        // link.setAttribute("href", csvUrl);
        link.setAttribute("href", uri);
        link.click();
      }
      else {
        if (window.navigator.msSaveOrOpenBlob) { //IE>=10
          link.addEventListener('click', function () {
            window.navigator.msSaveOrOpenBlob(blob, filename);
          });
          link.click();


        } else { //支持IE9、IE8;  IE7及以下暂不支持，因为不支持JSON
          var ifd = document.getElementById('ifdiframe').contentDocument;
          ifd.open('text/plain', 'replace');
          ifd.write('\r\n\r\n' + csv);
          ifd.close();
          ifd.execCommand('SaveAs', null, filename);
        }
      }

      document.body.removeChild(link);
      document.body.removeChild(ifdo);
    }
    //自定义字段数据
    var customFields = [];
    const userId = Meteor.userId();
    const params = {
      boardId: Session.get('currentBoard'),
    };
    const queryParams = {
      authToken: Accounts._storedLoginToken(),
    };
    //根据id替换姓名
    function changeName(num, arr) {
      if (num == undefined) {
        return "未选择人员";
      }
      for (var i in arr) {
        if (arr[i]["_id"] == num) {
          return arr[i]["profile"]["fullname"];
        }
      }
    }
    function getcont(data, name) {
      let dataFilter = [], custom = [];
      custom = customFields.filter(item => item.name == name);
      if (custom.length > 0) {
        dataFilter = data.filter(item => item._id == custom[0]._id)
      }
      if (dataFilter.length > 0 && dataFilter[0].value) {
        return dataFilter[0].value;
      } else {
        return "未填写" + name;
      }

    }
    //格式化时间
    function replaceTime(date) {
      date = new Date(date);
      var y = date.getFullYear();
      if (y.toString() == "NaN") {
        return '未设置时间';
      } else {
        var m = date.getMonth() + 1;
        m = m < 10 ? ('0' + m) : m;
        var d = date.getDate();
        d = d < 10 ? ('0' + d) : d;
        var h = date.getHours();
        h = h < 10 ? ('0' + h) : h;
        var minute = date.getMinutes();
        minute = minute < 10 ? ('0' + minute) : minute;
        return y + '-' + m + '-' + d + ' ' + h + ':' + minute;
      }
    }

    //在看板活动中查找卡片从进行中移动到已完成-Release时的时间就是实际完成时间
    function getFactFinishTime(cardId, activities, name) {
      let filterData = activities.filter(item => item.activityType == "moveCard" && item.cardId == cardId && item.listName == name)
      return filterData[0] ? replaceTime(filterData[0].createdAt) : "未完成";
    }
    //获取标签
    function getxhLabelId(cardlabel, arr) {
      var lsarr = "";
      if (cardlabel.length == 0) {
        return "暂无标签";
      } else {
        for (var i in cardlabel) {
          lsarr += getlabel(cardlabel[i], arr) + ",";
        }
        lsarr = lsarr.substring(0, lsarr.length - 1);
        return lsarr;
      }
    }
    function getlabel(num, arr) {
      for (var i in arr) {
        if (arr[i]["_id"] == num) {
          return arr[i]["name"];
        }
      }
    }
    //获取状态
    function getstate(num, arr) {
      for (var i in arr) {
        if (arr[i]["_id"] == num) {
          return arr[i]["title"];
        }
      }
    }
    //根据选择项，将得到的选项放入数组
    function listArr(card, allData, id) {
      for (var i in allData["cards"]) {
        if (allData["cards"][i]["_id"] == id) {
          card.push(allData["cards"][i]);
        }
      }
    }
    //获取数据
    HTTP.get(FlowRouter.path('/api/boards/:boardId/custom-fields', params, queryParams), {}, function (error, response) {
      if (error) {
        console.log(error);
      } else {
        customFields = response.data;
      }
    });
    //获取数据
    HTTP.get(FlowRouter.path('/api/boards/:boardId/export', params, queryParams), {}, function (error, response) {
      if (error) {
        console.log(error);
      } else {
        var jsonData = [];
        jsonData.push(response.data);
        // 选中项放入数组
        var dataArr = [];
        var listCard = $(".list-body .minicards .minicard-wrapper");
        var lscards = [];
        for (var i in listCard) {
          if (!isNaN(i)) {
            if (listCard[i]["classList"]["value"].indexOf("is-checked") != "-1") {
              listArr(lscards, jsonData[0], listCard[i]["pathname"].substring(listCard[i]["pathname"].lastIndexOf("/") + 1, listCard[i]["pathname"].length));
            }
          }
        }
        if (lscards.length == 0) {
          lscards = jsonData[0]["cards"];
        }
        var n = 0;
        for (var i in lscards) {
          if (lscards[i]["archived"] == false) {
            n = n + 1;
            var jsons = {
              "序号": n,
              "工作名称": lscards[i]["title"],
              "工作内容": lscards[i]["description"],
              "本周进度": getcont(lscards[i]["customFields"], "本周进度"),
              "下周计划": getcont(lscards[i]["customFields"], "下周计划"),
              "交付物": getcont(lscards[i]["customFields"], "交付物"),
              "状态": getstate(lscards[i]["listId"], jsonData[0]["lists"]),
              "标签": getxhLabelId(lscards[i]["labelIds"], jsonData[0]["labels"]),
              "计划完成时间": replaceTime(lscards[i]["dueAt"]),
              "实际完成时间": getFactFinishTime(lscards[i]._id, jsonData[0]["activities"], "已完成"),
              "责任人员": changeName(lscards[i]["members"][0], jsonData[0]["users"]),
              "创建时间": replaceTime(lscards[i]["createdAt"])
            };
            dataArr.push(jsons)
          }

        }

        var GridData = JSON.parse(JSON.stringify(dataArr))
        JSONToCSV(GridData, "WeeklyReport.csv", true);
      }

    });
  },
  // 打印
  'click .printCards': function () {
    preview(1);
    function preview(oper) {
      if (oper < 10) {
        var content = document.getElementById("content");
        var header = document.getElementById("header");
        var header_quick_access = document.getElementById("header-quick-access");
        var pop_over = document.getElementById("pop-over");
        var tables = document.getElementById("tables");
        content.style.display = "none";
        header.style.display = "none";
        header_quick_access.style.display = "none";
        pop_over.style.display = "none";
        document.body.style.backgroundColor = "#fff";
        tables.style.display = "block";
        window.print();
        content.style.display = "block";
        header.style.display = "block";
        header_quick_access.style.display = "block";
        pop_over.style.display = "block";
        document.body.style.backgroundColor = "#DEDEDE";
        tables.style.display = "none";
      } else {
        window.print();
      }
    }
  },
});
