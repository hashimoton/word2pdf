# coding: utf-8
#
# word2pdf.rb
#   Saves Microsoft Word files as PDF files.
#   Usage: ruby word2pdf.rb word_file.doc [another_word_file.docx ...]

require "win32ole"

def absolute_path(file_path)
  fso_obj = WIN32OLE.new("Scripting.FileSystemObject")
  return fso_obj.GetAbsolutePathName(file_path)
end

def word2pdf(word, word_file)

  if word_file !~ /\.docx?$/
    return
  end

  word_path = absolute_path(word_file)
  pdf_path = word_path.sub(/\.docx?$/, ".pdf")

  doc = word.Documents.Open(word_path,
    {
      "ReadOnly" => true
    })

  doc.ExportAsFixedFormat(
    {
      "OutputFileName" => pdf_path,
      "ExportFormat" => 17,
      "OpenAfterExport" => true
    })

  doc.Close(
    {
       "SaveChanges" => false
    })
    
end


########################################

word_files = ARGV
word = WIN32OLE.new("Word.Application")

word_files.each do |word_file|
  word2pdf(word, word_file)
end

word.Quit

# EOF
