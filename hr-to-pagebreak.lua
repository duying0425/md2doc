function HorizontalRule(el)
  if FORMAT:match('latex') then
    return pandoc.RawBlock('latex', '\\newpage')
  elseif FORMAT:match('html') then
    return pandoc.RawBlock('html', '<div style="page-break-after: always;"></div>')
  elseif FORMAT:match('docx') then
    return pandoc.RawBlock('openxml', '<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
  else
    return el
  end
end