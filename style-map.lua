function Header(elem)
  print("Processing Header")
  if elem.classes:includes("csp-chapter-title") then
    elem.attributes["custom-style"] = "CSP - Chapter Title"
  end
  return elem
end