require "word_handler/version"
require 'win32ole'
require 'singleton'  

module WordHandler
  # Your code goes here...
  class Word
    
    include Singleton 
   
    attr_accessor :exe
    
    def initialize
      @exe = WIN32OLE.new('word.application')
      @exe.visible = false
    end
    
    def close
      @exe.quit
    end
    
    def givedoc(fpath = nil)
      fpath.nil? ?  @exe.Documents.Add() : @exe.Documents.Open("#{fpath}") ;
      
    end

  end


  class DocWriter

    def initialize(doc, name)
      @doc  = doc
      @name = name
  
    end
    
    def msg(message)

    
	  # 将当前文档设为活动状态
	  @doc.Activate 
	  @doc.Content.Font.Size = 11
	  @doc.Content.Text = "#{Time.now}: #{@name}: #{message}"

	  @doc.Save
	  @doc.close
      
    end

  end

=begin
  class LoggerFactory
    def initialize(bdir)
      @basedir = bdir
      @loggers = {}
    end
    
    def get_logger(name)
      if !@loggers.has_key? name
      
        fname = name.gsub(/[.\/]/, "_").untaint
        @loggers[name] = Logger.new(name, @basedir + "/" + fname)

      # 在word中，添加一个文档
	  doc = @wordexe.Documents.Add()
      end
      return @loggers[name]
    end

  end
=end

end
