   M S B A M L   `   `   `  ����?  ;Main, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:MainMain  5   	       Main.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   $localclr-namespace:Main     5   	   -   5          � Title$  
MainWindow��5   	   $	��500��6   $	��710��6*    � WindowStartupLocation$ CenterScreen=�66    � 
ResizeMode$ CanMinimize=�6[   .��5	      � .+�5
   
   ��  �� Orientation$ Vertical=�6   .+�5      �� $��指定 Word 檔案��6   $
��Left=�6+   $	��Top=�6F   $��	10,10,0,0q�6^   $��65��6q   $	��645��6}   .��5      �� $ 
Horizontal=�6   .+�5      �� -   5        $��TargetWordFilePathTextBox��5      $
��Left=�6@   $��23��6[   $��	10,10,0,0q�6g    �� TextWrapping$
 Wrap=�6z   $	��Top=�6�   $	��536��6�   �� 5      -   5        $��選取檔案��5      $
��Left=�6,   $��5,10,0,0q�6G   $	��Top=�6Y   $��75��6q    J� RenderTransformOrigin$ 1.247,-0.265&�6|   6�   5      5      �� 5      $ 
Horizontal=�6   .+�5      � $	��370��6   $	��200��6$   .+�5      R� -   5         R�Name$ foldersItem��5      $
��Auto��6j   ���	#FFFFFFFF���	#FFFFFFFF���	#FFFFFFFF � 	Resources 5       �� 
TargetType	 �� Property
 �� Value �� RelativeSource �� Mode �� AncestorType!   Main.HeaderToImageConverter   Instance �� 	Converter��%W  'Q�       �� 5      " Q�6$   .�5   "   �� 	 � ��
 5   &   �� 5   *   .R�5   .   �� $ 
Horizontal=�69   .+�5   2   �� $	��img��68   $��20��6D   $��20��6O   $
��Fill=�6[   ��5   8   ��  �� $ FindAncestor=�" Q�# Z �� 5   2   $	��5,0q�6M   ��6<   �� 6M   5   /   5    +   5"   #   5#      5$      5%      5&      5'      5(      5)      5*      