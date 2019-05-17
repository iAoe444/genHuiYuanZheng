# genHuiYuanZheng

> 用于批量生成物联网协会会员证，书写于2019/05/17 17:25

## 注意📌

* 在运行程序前确保是否已经安装好了`numpy`，`pystrich`，`opencv`，`PIL`，`xlrd`库，为了方便，我这里使用的是`anaconda`来安装这些库
* 会员相片的大小需要设置为正方形的，尺寸最好设置为312*312，格式为`png`
* 生成的会员证已经上下左右留了0.2cm的出血
* 由于分辨率问题，所以生成的会员证大小太大了，需要到打印店改为10*7.1的大小
* 会员证采用的单面折叠，成双面的方式，需要注意打印的时候一个正面在上，一张180度旋转的反面在下

## 使用说明📋

### 1. 导入人员信息

​	在`人员信息.xls`文件中，顶部输入年份，在姓名和编号一栏中输入会员姓名和编号，适当使用excel的快捷操作可以快速生成编号

![导入人员信息](https://ws1.sinaimg.cn/large/006bBmqIgy1g34gfx925hg30r00n1tf8.gif)

### 2.  导入会员相片

​	在`相片`文件夹中导入带会员姓名的相片，会员相片的大小需要设置为正方形的，尺寸最好设置为312*312，格式为`png`，如果会员没有照片可以不用放，会自动用会徽代替

![导入会员相片](https://ws1.sinaimg.cn/large/006bBmqIgy1g34gi5sef6g30r00n1djq.gif)

### 3 . 执行程序

​	在执行前确保需要安装的python库已经安装了，然后执行程序：`python genHuiYuanZheng.py`，在`front`文件夹中存放的是正面的相片，在`back`文件夹中存放的是背面的相片

![执行程序](https://ws1.sinaimg.cn/large/006bBmqIgy1g34glkjw98g30r00n14b7.gif)

## 需要改进的地方🐞

​	由于只是为了方便，写程序写的比较块，所以目前存在一些问题，希望有空的人帮忙改进，目前已发现的问题有:

- [ ] 文字居中问题——用了判断姓名字数的方法分三种情况来实现居中，不算是一种好的解决方法，可以进行改进
- [ ] 图片分辨率问题——需要把生成的图像的分辨率换成300，这样图片的尺寸才是对的
- [ ] 切割圆形问题——在切割圆形的时候，发现用算法切出来的方式图片边缘锯齿问题较为严重，所以用图片叠加的方式完成实现圆形的图片，算是比较讨巧的方式，可以对其进行改进