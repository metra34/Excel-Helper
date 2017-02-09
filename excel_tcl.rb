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
$separator = TkVariable.new

task_values = ['concatenate']

########################################################
# Create frame with content pane
######################################################## 
root = TkRoot.new {title "Excel Helper"}
content = Tk::Tile::Frame.new(root) {padding "3 3 12 12"}.grid( :sticky => 'nsew')
TkGrid.columnconfigure root, 0, :weight => 1; TkGrid.rowconfigure root, 0, :weight => 1

########################################################
# Create input fields and labels
######################################################## 
# # TODO progressbar
# p = Tk::Tile::Progressbar.new(parent) {orient 'horizontal'; length 200; mode 'indeterminate'}
# # TODO alert on complete
# Tk::messageBox :message => 'Have a good day'

# TODO fix browse filter
$file_path.value = ' SOMETHING HERE '
Tk::Tile::Label.new(content) {text 'Select File:'}.grid( :column => 1, :row => 1, :sticky => 'w')
Tk::Tile::Button.new(content) {text 'Browse'; command { $file_path.value = Tk::getOpenFile{ typevariable '.xls'} }}.grid( :column => 2, :row => 1, :sticky => 'w')

filepath = Tk::Tile::Label.new(content) {textvariable $file_path}.grid( :column => 1, :row => 2, :columnspan => 4, :sticky => 'we')

Tk::Tile::Label.new(content) {text 'Target File Name:'}.grid( :column => 1, :row => 3, :columnspan => 2, :sticky => 'we')
target_name = Tk::Tile::Entry.new(content) {width 10; textvariable $target_name}.grid( :column => 3, :row => 3, :columnspan => 2, :sticky => 'we' )

Tk::Tile::Label.new(content) {text 'Task:'}.grid( :column => 1, :row => 4, :sticky => 'w')
task = Tk::Tile::Combobox.new(content) { textvariable $task }.grid( :column => 3, :row => 4, :columnspan => 2, :sticky => 'we')
task.values = task_values

row = Tk::Tile::RadioButton.new(content) {text 'Row'; variable $direction; value 'row'}.grid( :column => 2, :row => 5, :sticky => 'e')
col = Tk::Tile::RadioButton.new(content) {text 'Column'; variable $direction; value 'col'}.grid( :column => 3, :row => 5, :sticky => 'e')

Tk::Tile::Label.new(content) {text 'Range:'}.grid( :column => 1, :row => 6, :sticky => 'w')

Tk::Tile::Label.new(content) {text 'Rows:'}.grid( :column => 1, :row => 7, :sticky => 'w')
Tk::Tile::Entry.new(content) {width 5; textvariable $row_to}.grid( :column => 2, :row => 7, :sticky => 'e' )
Tk::Tile::Label.new(content) {text ' to '}.grid( :column => 3, :row => 7, :sticky => 'e')
Tk::Tile::Entry.new(content) {width 5; textvariable $row_from}.grid( :column => 4, :row => 7, :sticky => 'w' )

Tk::Tile::Label.new(content) {text 'Columns:'}.grid( :column => 1, :row => 8, :sticky => 'w')
Tk::Tile::Entry.new(content) {width 5; textvariable $col_to}.grid( :column => 2, :row => 8, :sticky => 'e' )
Tk::Tile::Label.new(content) {text 'to'}.grid( :column => 3, :row => 8, :sticky => 'e')
Tk::Tile::Entry.new(content) {width 5; textvariable $col_from}.grid( :column => 4, :row => 8, :sticky => 'w' )

Tk::Tile::Label.new(content) {text 'Target:'}.grid( :column => 1, :row => 9, :sticky => 'w')
Tk::Tile::Entry.new(content) {width 5; textvariable $target_element}.grid( :column => 2, :row => 9, :sticky => 'e' )
Tk::Tile::Label.new(content) {text 'Separator:'}.grid( :column => 3, :row => 9, :sticky => 'e')
Tk::Tile::Entry.new(content) {width 3; textvariable $separator}.grid( :column => 4, :row => 9, :sticky => 'w' )

Tk::Tile::Button.new(content) {text 'Run'; command {calculate}}.grid( :column => 4, :row => 10, :sticky => 'w')

########################################################
# Add padding and bindings
######################################################## 
TkWinfo.children(content).each {|w| TkGrid.configure w, :padx => 5, :pady => 5}
# root.bind("Return") {calculate}

########################################################
# functions start
######################################################## 

def calculate
	begin
		$direction.value = 'hello'
	rescue
		$direction.value = ''
	end
end

########################################################
# run the app
######################################################## 

Tk.mainloop
