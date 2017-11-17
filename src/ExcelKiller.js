var fs = require('fs');
var xlsx = require("node-xlsx");
var request = require("request");
var mkdirp = require('mkdirp');
var Promise = require('promise');

var ExcelKiller = function(filename) {
    this.dir = '../images';
    this.list = filename ? xlsx.parse("../excel/" + filename + '.xlsx') : xlsx.parse("../excel/" + 'import.xlsx');
    this.findImgList = [];
    this.num = 0;

    //下载图片
    this.download = function(url, dir, filename) {
        var promise = new Promise(function(resolve, reject) {
            request.get(url).on('error', function(err) {
                console.log(err)
            }).on('response', function(response) {
                resolve();
            }).pipe(fs.createWriteStream(dir + "/" + filename));
        });
        return promise;
    }

    var _this = this;
    var srcStr;
    var srcStrTemp = ''

    //删除生成的目录
    function deleteFolder(path) {
        var files = [];
        if (fs.existsSync(path)) {
            files = fs.readdirSync(path);
            files.forEach(function(file, index) {
                var curPath = path + "/" + file;
                if (fs.statSync(curPath).isDirectory()) { // recurse
                    deleteFolder(curPath);
                } else { // delete file
                    fs.unlinkSync(curPath);
                }
            });
            fs.rmdirSync(path);
        }
    }

    //删除生成的文件
    function deleteFile(path) {
        if (fs.existsSync(path)) {
            fs.unlinkSync(path);
        }
    }

    //重置
    function reset() {
        deleteFolder('../images');
        deleteFile('../exportExcel.xlsx');
    }

    //获取时间戳
    function getTimestamp() {
        return Math.round(new Date().getTime());
    }

    //创建目录
    function makeDir() {
        mkdirp(_this.dir, function(err) {
            if (err) {
                console.log(err);
            }
        });
    }

    //整理出下载地址
    function makeDownloadUrl(imgUrl) {
        imgUrl = imgUrl.replace('src="https:', '');
        imgUrl = imgUrl.replace('src="', '');
        imgUrl = imgUrl.replace('"', '');
        imgUrl = 'http:' + imgUrl;
        return imgUrl;
    };

    //写出image名字并且整理出img标签
    function makeImageName(extName) {
        return getTimestamp().toString() + parseInt(Math.random() * 1000000).toString() + extName;
    };

    reset();
    setTimeout(function(){
        makeDir();
        for (var i = 1; i < _this.list[0].data.length; i++) {
            var findImg = _this.list[0].data[i][2].match(/src=\"https:\/\/.*?\"|src=\"\/\/.*?\"/g);
            for (var j = 0; j < findImg.length; j++) {
                var imgUrl = makeDownloadUrl(findImg[j])
                var imageName = makeImageName(imgUrl.substr(-4, 4))
                srcStr = '<img src="images/' + imageName + '">';
                srcStrTemp += srcStr
                _this.findImgList.push([imgUrl, imageName]);
            }
            _this.list[0].data[i].push(srcStrTemp);
            srcStrTemp = '';
        }
    },2000)
}

//写excel文件
ExcelKiller.prototype.writeExcel = function() {
    var data = this.list[0].data;
    var buffer = xlsx.build([{ name: "mySheetName", data: data }]); // Returns a buffer
    fs.appendFile('../exportExcel.xlsx', buffer);
};

//下载图片
ExcelKiller.prototype.downloadImages = function() {
    if (this.findImgList.length !== 0) {
        var _this = this;
        this.download(this.findImgList[this.num][0], this.dir, this.findImgList[this.num][1]).then(function() {
            if (_this.num === _this.findImgList.length) {
                _this.num = 0;
                return;
            } else {
                _this.num++;
                _this.downloadImages();
            }
        });
    }
};

module.exports = ExcelKiller;