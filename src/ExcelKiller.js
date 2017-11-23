var fs = require('fs');
var path = require('path');
var xlsx = require("node-xlsx");
var request = require("request");
var mkdirp = require('mkdirp');
var Promise = require('promise');

var ExcelKiller = function(filename) {
    this.dir = [path.join(__dirname, '../imagesFromDetail'), path.join(__dirname, '../imagesFromList')];
    this.export = path.join(__dirname, '../exportExcel.xlsx');
    this.list = filename ? xlsx.parse(path.join(__dirname, '../excel/') + filename + '.xlsx') : xlsx.parse(path.join(__dirname, '../excel/') + 'import.xlsx');
    this.findImgList = [];
    this.howManyList = 0;
    this.count = [];

    //下载图片
    this.download = function(url, dir, filename) {
        var promise = new Promise(function(resolve, reject) {
            request.get(url).on('error', function(err) {
                console.log(err)
                resolve();
            }).on('response', function(response) {
                resolve();
            }).pipe(fs.createWriteStream(dir + "/" + filename));
        });
        return promise;
    }

    var _this = this;

    //删除生成的目录
    function deleteFolder(_path) {
        var files = [];
        if (fs.existsSync(_path)) {
            files = fs.readdirSync(_path);
            files.forEach(function(file, index) {
                var curPath = _path + "/" + file;
                if (fs.statSync(curPath).isDirectory()) { // recurse
                    deleteFolder(curPath);
                } else { // delete file
                    fs.unlinkSync(curPath);
                }
            });
            fs.rmdirSync(_path);
        }
    }

    //删除生成的文件
    function deleteFile(_path) {
        if (fs.existsSync(_path)) {
            fs.unlinkSync(_path);
        }
    }

    //重置
    function reset() {
        for (var i = 0; i < _this.dir.length; i++) {
            deleteFolder(_this.dir[i]);
        }
        deleteFile(_this.export);
    }

    //获取时间戳
    function getTimestamp() {
        return Math.round(new Date().getTime());
    }

    //创建目录
    function makeDir() {
        for (var i = 0; i < _this.dir.length; i++) {
            mkdirp(_this.dir[i], function(err) {
                if (err) {
                    console.log(err);
                }
            });
        }
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

    function assetArray(num) {
        _this.howManyList++;
        _this.count.push(0);
        var array = [];
        _this.findImgList.push(array);
        var srcStr;
        var srcStrTemp = '';
        for (var i = 1; i < _this.list[0].data.length; i++) {
            var findImg = _this.list[0].data[i][num].match(/src=\"https:\/\/.*?\"|src=\"\/\/.*?\"/g);
            if (findImg !== null) {
                for (var j = 0; j < findImg.length; j++) {
                    var imgUrl = makeDownloadUrl(findImg[j])
                    var imageName = makeImageName(imgUrl.substr(-4, 4))
                    srcStr = '<img src="images/' + imageName + '">';
                    srcStrTemp += srcStr
                    _this.findImgList[_this.howManyList - 1].push([imgUrl, imageName]);
                }
                _this.list[0].data[i].push(srcStrTemp);
                srcStrTemp = '';
            }
        }
    }

    reset();
    setTimeout(function() {
        makeDir();
        assetArray(3);
        assetArray(4);
    }, 2000);
}

//写excel文件
ExcelKiller.prototype.writeExcel = function() {
    var data = this.list[0].data;
    var buffer = xlsx.build([{ name: "mySheetName", data: data }]); // Returns a buffer
    fs.appendFile(this.export, buffer);
};

//下载图片
ExcelKiller.prototype.downloadImages = function(num) {
    if (this.findImgList.length !== 0) {
        var _this = this;
        if (this.findImgList[num].length !== 0) {
            this.download(this.findImgList[num][this.count[num]][0], this.dir[num], this.findImgList[num][this.count[num]][1]).then(function() {
                if (_this.count[num] === _this.findImgList[num].length) {
                    _this.count[num] = 0;
                    return;
                } else {
                    _this.count[num]++;
                    _this.downloadImages(num);
                }
            });
        }
    }
};

module.exports = ExcelKiller;