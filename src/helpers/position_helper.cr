def int_pos_to_char(x : Int32 | Int64, y : Int32 | Int64)
  raise "int_pos_to_char method doesn't allow x > 25 for the moment" if x > 25
  (x+65).unsafe_chr+(y+1).to_s
end
