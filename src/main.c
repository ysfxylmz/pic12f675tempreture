#include <12F675.h>


#FUSES NOWDT                    //Watch Dog Timer'ý kapattýk. (Kapatmasak pic döngü uzadýðýnda reset atar.)
#FUSES INTRC_IO                 //(Dýþarýdan bir osilatör baðlamak yerine entegre içindeki RC osilatör kullanýldý.)
#FUSES PUT                      //Power Up Timer (Pic çalýþtýðýnda dengeye ulaþana kadar bekler.)
#FUSES NOBROWNOUT               //
#FUSES NOMCLR                   //MCLR bacaðýný iptal edip yerine bi buton baðlandý.

//**************PORTLARA ÝSÝM VERÝLDÝ
#define LCD_DATA_PIN PIN_A1
#define LCD_CLOCK_PIN PIN_A0
#define LCD_EN_PIN PIN_A2
#define LCD_LT_PIN PIN_A5
#define LM35 sAN3
#define RLED PIN_A0
#define GLED PIN_A1
#define BLED PIN_A2
//*********************************



setup_oscillator( OSC_4MHZ );//RC osilatör 4MHZ ayarlandý. 

#use delay(clock=4MHz)//delay fonksiyonlarý 4mhz ile çalýþacak.
#include <3WireLCD.c>//LCD'nin 3 pinle çalýþmasý için kütüphane


//*******global deðiþkenler tanýmlandý
unsigned int8 sicaklik;
const float lsb=5.0/1023.0;
//**************************



void lcd_update()//lcdyi güncelleyen fonksiyon.
{

  lcd_initialize();                     // Initialize LCD 
  lcd_cmd(LCD_CLEAR);                    // ekraný temizler
  lcd_goto(1, 1);                        // satýr sütun seçme
  printf(lcd_out, "%3u",sicaklik);//yazdýr
  

}


void analog_read()//analog okuma
{

   set_adc_channel(3);//3 nolu adc pini ayarlandý.
   delay_ms(10);//delay
   sicaklik = read_adc(ADC_START_AND_READ);//oku ve sýcaklýða ata 
   sicaklik=sicaklik*lsb*100.0*2;//sýcaklýk hesabý
  
   
}





void RGB_ON(unsigned int8 tR,unsigned int8 tG, unsigned int8 tB)
//Bu fonksiyonda tr tg tb deðerleri kadar  istenilen porttan high verilir. 
//(eðer deðerler 255ten az ise 255 e tamamlanana kadar istenilen porttan low verilir)
{
unsigned int8 a;
unsigned int8 i;

for(a=0;a<25;a++)
{
   for(i=1;i<255;i++){  // 255 adymda 3 ayri PWM isaret uret
      if(i<=tR)output_high(RLED);//RLED=1;
      if(i>tR)output_low(RLED);//RLED=0;
   
      if(i<=tG)output_high(GLED);//GLED=1;
      if(i>tG)output_low(GLED);//GLED=0;
      
      if(i<=tB)output_high(BLED);//BLED=1;
      if(i>tB)output_low(BLED);//BLED=0; 
   
      delay_us(10); 
                  }
}
}






void bak_renk_tablosu(){
//sýcaklýða göre tr tg tb deðerleri 0-255 arasýnda belirlenir ve RGB_on fonksiyonuna gider.


if(sicaklik>=20 && sicaklik<21)
{
RGB_ON(255,0,255);
} // pembe

if(sicaklik>=21 && sicaklik<22)
{
RGB_ON(204,0,255);
}

if(sicaklik>=22 && sicaklik<23)
{
RGB_ON(153,0,255);
}

if(sicaklik>=23 && sicaklik<24)
{
RGB_ON(102,0,255);
}

if(sicaklik>=24 && sicaklik<25)
{
RGB_ON(51,0,255);
}

if(sicaklik>=25 && sicaklik<26)
{
RGB_ON(0,0,255);
} // mavi

if(sicaklik>=21 && sicaklik<24)
{
RGB_ON(0,51,255);
}



if(sicaklik>=24 && sicaklik<27)
{
RGB_ON(0,153,255);
}



if(sicaklik>=27 && sicaklik<30)
{
RGB_ON(0,255,255);
} // turkuaz



if(sicaklik>=30 && sicaklik<33)
{
RGB_ON(0,255,153);
}



if(sicaklik>=33 && sicaklik<36)
{
RGB_ON(0,255,0);
} // yesil


if(sicaklik>=37 && sicaklik<41)
{
RGB_ON(102,255,0);
}



if(sicaklik>=41 && sicaklik<42)
{
RGB_ON(255,255,0);
} // sari

if(sicaklik>=42 && sicaklik<43)
{
RGB_ON(255,153,0);
}

if(sicaklik>=43 && sicaklik<44)
{
RGB_ON(255,102,0);
}



if(sicaklik>=45)
{
RGB_ON(255,0,0);
} 

}


void main()
{
   setup_adc(adc_clock_div_32);//adc ölçüm oraný
   setup_adc_ports(LM35); //hangi porttan adc ölçüm alacaðý belirlendi
    
   while(TRUE)//sonsuz loop
   {
      analog_read();//okuma yapan fonksiyona gider
      if(INPUT(PIN_A3) == FALSE)//eðer buton 0'da ise lcd ye yazar deðilse sadece led rengi deðiþir
      {
         lcd_update();
         delay_ms(250);
      }
      else
      {
      output_low(LCD_EN_PIN);
      bak_renk_tablosu();
      }     
      }

 
}

