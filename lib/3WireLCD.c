#define LCD_FIRST_ROW          0x80
#define LCD_SECOND_ROW         0xC0
#define LCD_THIRD_ROW          0x94
#define LCD_FOURTH_ROW         0xD4
#define LCD_CLEAR              0x01
#define LCD_RETURN_HOME        0x02
#define LCD_CURSOR_OFF         0x0C
#define LCD_UNDERLINE_ON       0x0E
#define LCD_BLINK_CURSOR_ON    0x0F
#define LCD_MOVE_CURSOR_LEFT   0x10
#define LCD_MOVE_CURSOR_RIGHT  0x14
#define LCD_TURN_ON            0x0C
#define LCD_TURN_OFF           0x08
#define LCD_SHIFT_LEFT         0x18
#define LCD_SHIFT_RIGHT        0x1E

short RS;
void lcd_write_nibble(unsigned int8 n){
  unsigned int8 i;
  output_low(LCD_CLOCK_PIN);
  output_low(LCD_EN_PIN);
  for( i = 8; i > 0; i = i >> 1){
    if(n & i)
      output_high(LCD_DATA_PIN);
    else
      output_low(LCD_DATA_PIN);
    Delay_us(10);
    output_high(LCD_LT_PIN);
    output_high(LCD_clock_PIN);
    
    Delay_us(10);
    output_low(LCD_LT_PIN);
    output_low(LCD_clock_PIN);
    
  }
  if(RS)
    output_high(LCD_DATA_PIN);
  else
    output_low(LCD_DATA_PIN);
  for(i = 0; i < 2; i++){
    Delay_us(10);
    output_high(LCD_LT_PIN);
    output_high(LCD_clock_PIN);
    
    Delay_us(10);
    output_low(LCD_LT_PIN);
    output_low(LCD_clock_PIN);
    
  }
  output_high(LCD_EN_PIN);
  delay_us(2);
  output_low(LCD_EN_PIN);
}

void LCD_Cmd(unsigned int8 Command){
  RS = 0;
  lcd_write_nibble(Command >> 4);
  lcd_write_nibble(Command & 0x0F);
  if((Command == 0x0C) || (Command == 0x01) || (Command == 0x0E) || (Command == 0x0F)
  || (Command == 0x10) || (Command == 0x1E) || (Command == 0x18) || (Command == 0x08)
  || (Command == 0x14) || (Command == 0x02))
    Delay_ms(50);
}

void LCD_GOTO(unsigned int8 col, unsigned int8 row){
  switch(row){
    case 1:
      LCD_Cmd(0x80 + col-1);
      break;
    case 2:
      LCD_Cmd(0xC0 + col-1);
      break;
    case 3:
      LCD_Cmd(0x94 + col-1);
      break;
    case 4:
      LCD_Cmd(0xD4 + col-1);
    break;
  }
}

void LCD_Out(unsigned int8 LCD_Char){
  RS = 1;  
  lcd_write_nibble(LCD_Char >> 4);
  delay_us(10);
  lcd_write_nibble(LCD_Char & 0x0F);break;
}

void LCD_Initialize(){
  RS = 0;
  output_low(LCD_DATA_PIN);
  output_low(LCD_CLOCK_PIN);
  output_low(LCD_EN_PIN);
  output_low(LCD_LT_PIN);
  output_drive(LCD_DATA_PIN);
  output_drive(LCD_CLOCK_PIN);
  output_drive(LCD_EN_PIN);
  output_drive(LCD_LT_PIN);
  delay_ms(40);
  lcd_Cmd(3);
  delay_ms(5);
  lcd_Cmd(3);
  delay_ms(5);
  lcd_Cmd(3);
  delay_ms(5);
  lcd_Cmd(2);
  delay_ms(5);
  lcd_Cmd(0x28);
  delay_ms(50);
  lcd_Cmd(0x0C);
  delay_ms(50);
  lcd_Cmd(0x06);
  delay_ms(50);
  lcd_Cmd(0x0C);
  delay_ms(50);
}
