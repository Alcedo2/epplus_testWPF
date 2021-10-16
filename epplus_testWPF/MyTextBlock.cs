using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace epplus_testWPF
{
    /// <summary>
    /// このカスタム コントロールを XAML ファイルで使用するには、手順 1a または 1b の後、手順 2 に従います。
    ///
    /// 手順 1a) 現在のプロジェクトに存在する XAML ファイルでこのカスタム コントロールを使用する場合
    /// この XmlNamespace 属性を使用場所であるマークアップ ファイルのルート要素に
    /// 追加します:
    ///
    ///     xmlns:MyNamespace="clr-namespace:epplus_testWPF"
    ///
    ///
    /// 手順 1b) 異なるプロジェクトに存在する XAML ファイルでこのカスタム コントロールを使用する場合
    /// この XmlNamespace 属性を使用場所であるマークアップ ファイルのルート要素に
    /// 追加します:
    ///
    ///     xmlns:MyNamespace="clr-namespace:epplus_testWPF;assembly=epplus_testWPF"
    ///
    /// また、XAML ファイルのあるプロジェクトからこのプロジェクトへのプロジェクト参照を追加し、
    /// リビルドして、コンパイル エラーを防ぐ必要があります:
    ///
    ///     ソリューション エクスプローラーで対象のプロジェクトを右クリックし、
    ///     [参照の追加] の [プロジェクト] を選択してから、このプロジェクトを参照し、選択します。
    ///
    ///
    /// 手順 2)
    /// コントロールを XAML ファイルで使用します。
    ///
    ///     <MyNamespace:MyTextBlock/>
    ///
    /// </summary>
    public class MyTextBlock : TextBlock
    {
        #region Properties

        public new string Text
        {
            get { return (string)GetValue(TextProperty); }
            set { SetValue(TextProperty, value); }
        }

        public new static readonly DependencyProperty TextProperty =
            DependencyProperty.Register("Text", typeof(string),
            typeof(MyTextBlock), new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.AffectsRender,
                new PropertyChangedCallback(UpdateHighlighting)));

        public RichTextContainer RichText
        {
            get { return (RichTextContainer)GetValue(RichTextProperty); }
            set { SetValue(RichTextProperty, value); }
        }

        public static readonly DependencyProperty RichTextProperty =
            DependencyProperty.Register("RichText", typeof(RichTextContainer),
            typeof(MyTextBlock), new FrameworkPropertyMetadata(new RichTextContainer(), FrameworkPropertyMetadataOptions.AffectsRender,
                new PropertyChangedCallback(UpdateHighlighting)));

        public string HighlightPhrase
        {
            get { return (string)GetValue(HighlightPhraseProperty); }
            set { SetValue(HighlightPhraseProperty, value); }
        }

        public static readonly DependencyProperty HighlightPhraseProperty =
            DependencyProperty.Register("HighlightPhrase", typeof(string),
            typeof(MyTextBlock), new FrameworkPropertyMetadata(String.Empty, FrameworkPropertyMetadataOptions.AffectsRender,
                new PropertyChangedCallback(UpdateHighlighting)));

        public Brush HighlightBrush
        {
            get { return (Brush)GetValue(HighlightBrushProperty); }
            set { SetValue(HighlightBrushProperty, value); }
        }

        public static readonly DependencyProperty HighlightBrushProperty =
            DependencyProperty.Register("HighlightBrush", typeof(Brush),
            typeof(MyTextBlock), new FrameworkPropertyMetadata(Brushes.Yellow, FrameworkPropertyMetadataOptions.AffectsRender,
                new PropertyChangedCallback(UpdateHighlighting)));

        public bool IsCaseSensitive
        {
            get { return (bool)GetValue(IsCaseSensitiveProperty); }
            set { SetValue(IsCaseSensitiveProperty, value); }
        }

        public static readonly DependencyProperty IsCaseSensitiveProperty =
            DependencyProperty.Register("IsCaseSensitive", typeof(bool),
            typeof(MyTextBlock), new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.AffectsRender,
                new PropertyChangedCallback(UpdateHighlighting)));

        private static void UpdateHighlighting(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            ApplyTextColor(d as MyTextBlock);
        }

        #endregion

        #region Members

        private static void ApplyTextColor(MyTextBlock tb)
        {
            if (tb.RichText == null || tb.RichText.getTextList() == null) return;
            tb.Inlines.Clear();
            foreach (MyFont tac in tb.RichText.getTextList())
            {
                if(tac.backGroundColor == null)
                    tb.Inlines.Add(new Run(tac.text)
                    {
                        Foreground = tac.foreGroundColor
                    });
                else
                    tb.Inlines.Add(new Run(tac.text)
                    {
                        Foreground = tac.foreGroundColor,
                        Background = tac.backGroundColor
                    });
            }
        }

        #endregion
    }

    public class RichTextContainer
    {
        List<MyFont> richTexts;

        public RichTextContainer()
        {
            richTexts = new List<MyFont>();
        }


        public void add(MyFont tac)
        {
            richTexts.Add(tac);
        }

        public void add(string str, Brush fc)
        {
            richTexts.Add(new MyFont(str, fc));
        }

        public void add(string str, Brush fc, Brush bc)
        {
            richTexts.Add(new MyFont(str, fc, bc));
        }

        public void clear()
        {
            richTexts.Clear();
        }

        public List<MyFont> getTextList()
        {
            return richTexts;
        }

    }

    public class MyFont 
    {
        public string text { get; set; }
        public Brush backGroundColor { get; set; }
        public Brush foreGroundColor { get; set; }

        public MyFont(string str, Brush fore)
        {
            text = str;
            backGroundColor = null;
            foreGroundColor = fore;
        }

        public MyFont(string str, Brush fore, Brush bk)
        {
            text = str;
            if(bk != null)
                backGroundColor = bk;
            foreGroundColor = fore;
        }
    }

}
