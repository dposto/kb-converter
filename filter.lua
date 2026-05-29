-- background-to-highlight-direct.lua
-- Applies highlighting formatting DIRECTLY instead of using custom styles
-- This works around pandoc stripping custom style definitions

function Span(el)
  -- Check if the span has a background-color style
  if el.attributes and el.attributes.style then
    local style = el.attributes.style
    
    -- Extract background-color value  
    local bgcolor = style:match('background%-color:%s*([^;]+)')
    
    if bgcolor then
      bgcolor = bgcolor:gsub('%s+', ''):lower()
      
      -- Map colors to their properties
      local color_props = {
        -- Color names
        yellow = {bg = "FFFF00", fg = "000000"},
        red = {bg = "FF0000", fg = "FFFFFF"},
        orange = {bg = "FFA500", fg = "000000"},
        black = {bg = "000000", fg = "FFFFFF"},
        
        -- Hex codes
        ["#ffff00"] = {bg = "FFFF00", fg = "000000"},
        ["#ff0"] = {bg = "FFFF00", fg = "000000"},
        ["#ff0000"] = {bg = "FF0000", fg = "FFFFFF"},
        ["#f00"] = {bg = "FF0000", fg = "FFFFFF"},
        ["#ffa500"] = {bg = "FFA500", fg = "000000"},
        ["#000000"] = {bg = "000000", fg = "FFFFFF"},
        ["#000"] = {bg = "000000", fg = "FFFFFF"},
        
        -- RGB format
        ["rgb(255,255,0)"] = {bg = "FFFF00", fg = "000000"},
        ["rgb(255,0,0)"] = {bg = "FF0000", fg = "FFFFFF"},
        ["rgb(255,165,0)"] = {bg = "FFA500", fg = "000000"},
        ["rgb(0,0,0)"] = {bg = "000000", fg = "FFFFFF"}
      }
      
      local props = color_props[bgcolor]
      if props then
        local padded_content = {}
        
        -- Check if content already has padding
        local first_elem = el.content[1]
        local starts_with_space = false
        if first_elem and first_elem.t == "Str" then
          if first_elem.text:match("^[\u{00A0} ]+") then
            starts_with_space = true
          end
        end
        
        local last_elem = el.content[#el.content]
        local ends_with_space = false
        if last_elem and last_elem.t == "Str" then
          if last_elem.text:match("[\u{00A0} ]+$") then
            ends_with_space = true
          end
        end
        
        -- Add padding if needed
        if not starts_with_space then
          table.insert(padded_content, pandoc.Str("\u{00A0}\u{00A0}"))
        end
        
        for _, item in ipairs(el.content) do
          table.insert(padded_content, item)
        end
        
        if not ends_with_space then
          table.insert(padded_content, pandoc.Str("\u{00A0}\u{00A0}"))
        else
          table.insert(padded_content, pandoc.Str(" "))
        end
        
        -- Return raw Word XML with complete run structure
        local content_text = pandoc.utils.stringify(pandoc.Span(padded_content))
        
        return pandoc.RawInline('openxml',
          '<w:r>' ..
          '<w:rPr>' ..
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>' ..
          '<w:b/>' ..
          '<w:color w:val="' .. props.fg .. '"/>' ..
          '<w:shd w:val="clear" w:fill="' .. props.bg .. '"/>' ..
          '<w:sz w:val="24"/>' ..
          '</w:rPr>' ..
          '<w:t xml:space="preserve">' .. content_text .. '</w:t>' ..
          '</w:r>')
      end
    end
  end
  
  return el
end

-- Style internal links
function Link(el)
  if el.target:match('^#') then
    -- Get link text
    local link_text = pandoc.utils.stringify(el)
    
    -- Return raw Word XML with complete run structure for 12pt Times New Roman blue underline
    return pandoc.RawInline('openxml',
      '<w:r>' ..
      '<w:rPr>' ..
      '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>' ..
      '<w:color w:val="0000FF"/>' ..
      '<w:u w:val="single"/>' ..
      '<w:sz w:val="24"/>' ..
      '</w:rPr>' ..
      '<w:t xml:space="preserve">' .. link_text .. '</w:t>' ..
      '</w:r>')
  end
  
  return el
end

-- Remove empty list items
local function has_content(blocks)
  for _, block in ipairs(blocks) do
    if block.t == "Para" or block.t == "Plain" then
      if #block.content > 0 then
        for _, inline in ipairs(block.content) do
          if inline.t == "Str" and inline.text:match('%S') then
            return true
          elseif inline.t ~= "Space" and inline.t ~= "SoftBreak" and inline.t ~= "LineBreak" then
            return true
          end
        end
      end
    elseif block.t ~= "Plain" and block.t ~= "Para" then
      return true
    end
  end
  return false
end

function BulletList(el)
  local new_items = {}
  for _, item in ipairs(el.content) do
    if has_content(item) then
      table.insert(new_items, item)
    end
  end
  el.content = new_items
  return el
end

function OrderedList(el)
  local new_items = {}
  for _, item in ipairs(el.content) do
    if has_content(item) then
      table.insert(new_items, item)
    end
  end
  el.content = new_items
  return el
end

-- Remove "AI Usage (No value)" section
function Pandoc(doc)
  local new_blocks = {}
  local skip_next_para = false
  
  for i, block in ipairs(doc.blocks) do
    if block.t == "Header" and block.level == 2 then
      local text = pandoc.utils.stringify(block)
      if text == "AI Usage" then
        skip_next_para = true
      else
        table.insert(new_blocks, block)
        skip_next_para = false
      end
    elseif skip_next_para and block.t == "Para" then
      local text = pandoc.utils.stringify(block)
      if text:match("%(No value%)") then
        skip_next_para = false
      else
        table.insert(new_blocks, block)
        skip_next_para = false
      end
    else
      table.insert(new_blocks, block)
    end
  end
  
  return pandoc.Pandoc(new_blocks, doc.meta)
end
