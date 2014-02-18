
require 'win32ole'
require 'singleton'  

module WordHandler

  class Word
    
    include Singleton 
   
    attr_accessor :exe
    @@doclist = {}

    def initialize
      @exe = WIN32OLE.new('wps.application')
      @exe.visible = false
    end
    
=begin
  功能：关闭word主程序
=end
    def close
      @exe.quit
    end
    
=begin
   创建doc文档
   参数：
        name    给doc取个名字
        fpath   doc文档的绝对路径 
=end
    def givedoc(name = nil, fpath = nil)
      unless FileTest::exist?(fpath)
        doc = @exe.Documents.Add()
        doc.Activate 
        doc.SaveAs("#{fpath}", 0)
      end
      doc = @exe.Documents.Open("#{fpath}") ;
      @@doclist["#{name}"] = doc
    end

=begin
   功能：关闭doc文档
   参数：
        name  给doc取的名字
=end
    def closedoc(name = nil)
      @@doclist["#{name}"].close
    end

=begin
    功能：往给定的doc文件添加信息
    参数：
          name    某个doc的名字
          message 输入的信息
=end
    def msg(name, message)
      # 将当前文档设为活动状态
      @@doclist["#{name}"].Activate 
      @@doclist["#{name}"].Content.Font.Size = 11
      #@@doclist["#{name}"].Content.Text = "#{Time.now}: #{message}"
      @@doclist["#{name}"].Range(@exe.Selection.End, @exe.Selection.End).Text = "#{Time.now}: #{message}\n"
      @@doclist["#{name}"].Save
      
    end

  end

end