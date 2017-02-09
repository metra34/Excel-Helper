########################################################
# requires ruby installed with tk interpreter to return
######################################################## 

require 'tk'
require 'tkextlib/tile'

########################################################
# Define input variables
######################################################## 
$file_path = TkVariable.new
$target_name = TkVariable.new
$task = TkVariable.new
$direction = TkVariable.new
$row_from = TkVariable.new
$row_to = TkVariable.new
$col_from = TkVariable.new
$col_to = TkVariable.new
$target_element = TkVariable.new

########################################################
# Create frame with content pane
######################################################## 
root = TkRoot.new {title "Excel Helper"}
content = Tk::Tile::Frame.new(root) {padding "3 3 12 12"}.grid( :sticky => 'nsew')
TkGrid.columnconfigure root, 0, :weight => 1; TkGrid.rowconfigure root, 0, :weight => 1

########################################################
# Create input fields and labels
######################################################## 
Tk::Tile::Label.new(content) {text 'Select File:'}.grid( :column => 1, :row => 1, :sticky => 'w')
# TODO file chooser

filepath = Tk::Tile::Label.new(content) {textvariable $file_path}.grid( :column => 1, :row => 2, :sticky => 'we')

Tk::Tile::Label.new(content) {text 'Target Name:'}.grid( :column => 1, :row => 3, :sticky => 'w')
target_name = Tk::Tile::Entry.new(content) {width 10; textvariable $target_name}.grid( :column => 2, :row => 3, :sticky => 'we' )

Tk::Tile::Label.new(content) {text 'Task:'}.grid( :column => 1, :row => 4, :sticky => 'w')
# TODO dropdown

row = Tk::Tile::RadioButton.new(content) {text 'Row'; variable $direction; value 'row'}.grid( :column => 1, :row => 5, :sticky => 'w')
col = Tk::Tile::RadioButton.new(content) {text 'Column'; variable $direction; value 'col'}.grid( :column => 2, :row => 5, :sticky => 'w')

Tk::Tile::Label.new(content) {text 'Range:'}.grid( :column => 1, :row => 6, :sticky => 'w')

Tk::Tile::Label.new(content) {text 'Rows:'}.grid( :column => 1, :row => 7, :sticky => 'w')
Tk::Tile::Entry.new(content) {width 3; textvariable $row_to}.grid( :column => 2, :row => 7, :sticky => 'we' )
Tk::Tile::Label.new(content) {text ' to '}.grid( :column => 3, :row => 7, :sticky => 'w')
Tk::Tile::Entry.new(content) {width 3; textvariable $row_from}.grid( :column => 4, :row => 7, :sticky => 'we' )

Tk::Tile::Label.new(content) {text 'Columns:'}.grid( :column => 1, :row => 8, :sticky => 'w')
Tk::Tile::Entry.new(content) {width 3; textvariable $col_to}.grid( :column => 2, :row => 8, :sticky => 'we' )
Tk::Tile::Label.new(content) {text ' to '}.grid( :column => 3, :row => 8, :sticky => 'w')
Tk::Tile::Entry.new(content) {width 3; textvariable $col_from}.grid( :column => 4, :row => 8, :sticky => 'we' )

Tk::Tile::Button.new(content) {text 'Execute'; command {calculate}}.grid( :column => 4, :row => 9, :sticky => 'w')

########################################################
# Add padding and bindings
######################################################## 
TkWinfo.children(content).each {|w| TkGrid.configure w, :padx => 5, :pady => 5}
f.focus
root.bind("Return") {calculate}

########################################################
# functions start
######################################################## 

def calculate
	begin
		$meters.value = (0.3048*$feet*10000.0).round()/10000.0
	rescue
		$meters.value = ''
	end
end

########################################################
# run the app
######################################################## 

Tk.mainloop
