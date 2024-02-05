require 'tk'
require 'win32ole'

$shell = WIN32OLE.new('WScript.Shell')


$average_word_length = 5

def calculate_delay()
    $words_per_minute = $speed_input.get('1.0','end').to_i
  if $words_per_minute<1
    $words_per_minute=45
  end
  characters_per_minute = ($words_per_minute * $average_word_length)
  characters_per_second = (characters_per_minute.to_f)/60.0
  delay = (1.0/(characters_per_second))
  delay.round(2)
  return delay
end

def autotype_sentence(text, delay = 0.25)
  text.chars.each do |char|
      if char=="\n"
        $shell.SendKeys("{Enter}")
      else
        $shell.SendKeys(char)
      end
      sleep(delay)
    end
end

def start_action
  input_text = $text_content.get('1.0','end')
  initial_delay = $initial_delay_input.get('1.0','end').to_i
  if(initial_delay<3 || initial_delay == nil)
      initial_delay = 3
  end
  sleep(initial_delay)
  delay = calculate_delay()
  autotype_sentence(input_text,delay)
end


root = TkRoot.new { title "Sammy's Little Helper" }

root.geometry('920x700')

$text_content = TkText.new(root) { height 100; width 80 }.pack(side: 'left', padx: 10, pady: 10)

start_button = TkButton.new(root) { text "Start"; command { start_action } }.pack(side: 'top', padx: 5, pady: 10)

speed_label = TkLabel.new(root) {text 'Speed (Default 45):'}.pack( padx: 10, pady: 10, side: 'top')

$speed_input = TkText.new(root) { height 1; width 10 }.pack(side: 'top', padx: 10, pady: 5)

delay_label = TkLabel.new(root) {text 'Delay (Default 3):'}.pack( padx: 10, pady: 10, side: 'top')

$initial_delay_input = TkText.new(root) { height 1; width 10 }.pack(side: 'top', padx: 10, pady: 5)



Tk.mainloop
