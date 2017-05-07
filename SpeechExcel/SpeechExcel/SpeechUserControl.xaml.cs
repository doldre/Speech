using System;
using System.Windows;
using System.Windows.Controls;
using SpeechExcel.Execute;

namespace SpeechExcel
{
    using Microsoft.CognitiveServices.SpeechRecognition;
    using System.ComponentModel;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;
    using System.IO.IsolatedStorage;
    using System.Runtime.CompilerServices;
    /// <summary>
    /// SpeechUserControl.xaml 的交互逻辑
    /// </summary>
    public partial class SpeechUserControl : UserControl, INotifyPropertyChanged
    {
        /// <summary>
        /// 可以把密钥放在app.config里面，不一定要从UI获取：string subscriptionKey = ConfigurationManager.AppSettings["primaryKey"];
        /// </summary>
        private string subscriptionKey;

        /// <summary>
        /// 数据识别region client
        /// </summary>
        private DataRecognitionClient dataClient;

        /// <summary>
        /// 麦克风客户端
        /// </summary>
        private MicrophoneRecognitionClient micClient;

        /// <summary>
        /// 这个对象控制语音信息的实时显示
        /// </summary>
        private string speakPartialInfo;

        /// <summary>
        /// Message Show in UI
        /// </summary>
        private string messageShow;

        /// <summary>
        /// 初始化新实例：<see cref="MainWindow"/>
        /// </summary>
        public SpeechUserControl()
        {
            this.InitializeComponent();
            this.Initialize();
        }

        #region Events

        /// <summary>
        /// Implement INotifyPropertyChanged interface
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        #endregion Events

        /// <summary>
        /// 获取/设置语音模式：长时音或者短时音
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is microphone client short phrase; otherwise, <c>false</c>.
        /// </value>
        public bool IsMicrophoneClientShortPhrase { get; set; }

        /// <summary>
        /// 获取/设置是否侦听麦克风信号，,涉及到长型号
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is microphone client dictation; otherwise, <c>false</c>.
        /// </value>
        public bool IsMicrophoneClientDictation { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is microphone client with intent.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is microphone client with intent; otherwise, <c>false</c>.
        /// </value>
        public bool IsMicrophoneClientWithIntent { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is data client short phrase.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is data client short phrase; otherwise, <c>false</c>.
        /// </value>
        public bool IsDataClientShortPhrase { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is data client with intent.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is data client with intent; otherwise, <c>false</c>.
        /// </value>
        public bool IsDataClientWithIntent { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is data client dictation.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is data client dictation; otherwise, <c>false</c>.
        /// </value>
        public bool IsDataClientDictation { get; set; }

        /// <summary>
        /// Gets or sets SubscriptionKey
        /// </summary>
        public string SubscriptionKey
        {
            get
            {
                return this.subscriptionKey;
            }
            set
            {
                this.subscriptionKey = value;
                this.OnPropertyChanged<string>();
            }
        }

        /// <summary>
        /// Gets or sets SpeekPartialContent
        /// </summary>
        public string SpeakPartialContent
        {
            get
            {
                return this.speakPartialInfo;
            }
            set
            {
                this.speakPartialInfo = value;
                this.OnPropertyChanged<string>("SpeakPartialContent");
            }
        }

        public string MessageShow
        {
            get
            {
                return this.messageShow;
            }
            set
            {
                this.messageShow = value;
                this.OnPropertyChanged<string>("MessageShow");
            }
        }

        public string ButtonStatus
        {
            get
            {
                return this._startbutton.Content.ToString();
            }
            set
            {
                this._startbutton.Content = value;
                this.OnPropertyChanged<string>("ButtonStatus");
            }
        }

        /// <summary>
        /// Gets the LUIS application identifier.
        /// </summary>
        /// <value>
        /// The LUIS application identifier.
        /// </value>
        private string LuisAppId
        {
            //get { return "00562bbb-2a3a-4e36-afa9-ccb398c7a103"; }
            get { return "1284fc06-1d2a-4dad-a49d-0aa0086af56c"; }
        }

        /// <summary>
        /// Gets the LUIS subscription identifier.
        /// </summary>
        /// <value>
        /// The LUIS subscription identifier.
        /// </value>
        private string LuisSubscriptionID
        {
            get { return "2f8b5ad36e6e44a2889702101e5e26bf"; }
        }

        /// <summary>
        /// 是否使用麦克风
        /// </summary>
        /// <value>
        ///   <c>true</c> if [use microphone]; otherwise, <c>false</c>.
        /// </value>
        private bool UseMicrophone
        {
            get
            {
                return this.IsMicrophoneClientWithIntent ||
                    this.IsMicrophoneClientDictation ||
                    this.IsMicrophoneClientShortPhrase;
            }
        }

        /// <summary>
        /// Gets a value indicating whether LUIS results are desired.
        /// </summary>
        /// <value>
        ///   <c>true</c> if LUIS results are to be returned otherwise, <c>false</c>.
        /// </value>
        private bool WantIntent
        {
            get
            {
                return false;
                //return !string.IsNullOrEmpty(this.LuisAppId) &&
                //    !string.IsNullOrEmpty(this.LuisSubscriptionID) &&
                //    (this.IsMicrophoneClientWithIntent || this.IsDataClientWithIntent);
            }
        }

        /// <summary>
        /// 获取当前的语音识别模式
        /// </summary>
        /// <value>
        /// The speech recognition mode.
        /// </value>
        /// speechclient的问题
        private SpeechRecognitionMode Mode
        {
            get
            {
                if (this.IsMicrophoneClientDictation ||
                    this.IsDataClientDictation)
                {
                    return SpeechRecognitionMode.LongDictation;
                }

                return SpeechRecognitionMode.ShortPhrase;
            }
        }

        /// <summary>
        /// Gets the default locale.
        /// </summary>
        /// <value>
        /// The default locale.
        /// </value>
        private string DefaultLocale
        {
            get { return "zh-CN"; }
        }

        /// <summary>
        /// Gets the Cognitive Service Authentication Uri.
        /// </summary>
        /// <value>
        /// The Cognitive Service Authentication Uri.  Empty if the global default is to be used.
        /// </value>
        private string AuthenticationUri
        {
            get
            {
                return ConfigurationManager.AppSettings["AuthenticationUri"];
            }
        }

        public bool CheckMic
        {
            get
            {
                return null != this.micClient;
            }
        }

        public void OnClose()
        {
            this.micClient.Dispose();
        }

        /// <summary>
        /// 初始化音频会话
        /// </summary>
        private void Initialize()
        {
            this.IsMicrophoneClientShortPhrase = true;
            // this setting for luis turn on
            this.IsMicrophoneClientWithIntent = true;
            // this setting for long tongue
            this.IsMicrophoneClientDictation = false;
            this.IsDataClientShortPhrase = false;
            this.IsDataClientWithIntent = false;
            this.IsDataClientDictation = false;
            this.SpeakPartialContent = "";
            this.done = true;
            this._startbutton.Content = "CLICK";

            //this.SubscriptionKey = ConfigurationManager.AppSettings["primarykey"];
            this.subscriptionKey = "7d41a160acf0421fa0c34a59659dc7f6";
        }

        /// <summary>
        /// 录音点击事件
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            // 禁用按钮，防止事件冲突
            this._startbutton.IsEnabled = false;
            this.MessageShow = "";
            ButtonStatus = "WAIT";

            if (this.UseMicrophone)
            {
                if (this.micClient == null)
                {
                    // 如果要实现luis请制定为true
                    if (this.WantIntent)
                    {
                        this.CreateMicrophoneRecoClientWithIntent();
                    }
                    else
                    {
                        // 不支持luis
                        this.CreateMicrophoneRecoClient();
                    }
                }
                // 启动麦克风并识别语音
                this.micClient.StartMicAndRecognition();
            }
            else
            {
                // 麦克风权限获取失败
                //MessageBox.Show("麦克风权限获取失败，请检查！");
                this.MessageShow = "麦克风权限获取失败，请检查!";
            }
        }

        /// <summary>
        /// Creates a new microphone reco client without LUIS intent support.
        /// </summary>
        private void CreateMicrophoneRecoClient()
        {
            this.micClient = SpeechRecognitionServiceFactory.CreateMicrophoneClient(
                this.Mode,
                this.DefaultLocale,
                this.SubscriptionKey);
            this.micClient.AuthenticationUri = this.AuthenticationUri;

            // Event handlers for speech recognition results
            this.micClient.OnMicrophoneStatus += this.OnMicrophoneStatus;
            this.micClient.OnPartialResponseReceived += this.OnPartialResponseReceivedHandler;
            if (this.Mode == SpeechRecognitionMode.ShortPhrase)
            {
                this.micClient.OnResponseReceived += this.OnMicShortPhraseResponseReceivedHandler;
            }
            else if (this.Mode == SpeechRecognitionMode.LongDictation)
            {
                this.micClient.OnResponseReceived += this.OnMicDictationResponseReceivedHandler;
            }

            this.micClient.OnConversationError += this.OnConversationErrorHandler;
        }

        /// <summary>
        /// Creates a new microphone reco client with LUIS intent support.
        /// </summary>
        private void CreateMicrophoneRecoClientWithIntent()
        {
            // this.WriteLine("--- Start microphone dictation with Intent detection ----");

            this.micClient =
                SpeechRecognitionServiceFactory.CreateMicrophoneClientWithIntent(
                this.DefaultLocale,
                this.SubscriptionKey,
                this.LuisAppId,
                this.LuisSubscriptionID);
            this.micClient.AuthenticationUri = this.AuthenticationUri;
            this.micClient.OnIntent += this.OnIntentHandler;

            // Event handlers for speech recognition results
            this.micClient.OnMicrophoneStatus += this.OnMicrophoneStatus;
            this.micClient.OnPartialResponseReceived += this.OnPartialResponseReceivedHandler;
            this.micClient.OnResponseReceived += this.OnMicShortPhraseResponseReceivedHandler;
            this.micClient.OnConversationError += this.OnConversationErrorHandler;
        }

        /// <summary>
        /// Called when a final response is received;
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechResponseEventArgs"/> instance containing the event data.</param>
        private void OnMicShortPhraseResponseReceivedHandler(object sender, SpeechResponseEventArgs e)
        {
            Dispatcher.Invoke((Action)(() =>
            {
                // this.WriteLine("--- OnMicShortPhraseResponseReceivedHandler ---");

                // we got the final result, so it we can end the mic reco.  No need to do this
                // for dataReco, since we already called endAudio() on it as soon as we were done
                // sending all the data.
                this.micClient.EndMicAndRecognition();
                this.WriteResponseResult(e);
                _startbutton.IsEnabled = true;
                if (e.PhraseResponse.RecognitionStatus == RecognitionStatus.InitialSilenceTimeout)
                {
                    //this.SpeakPartialContent = Properties.Resources.detect_warning;
                    this.MessageShow = Properties.Resources.detect_warning;
                }
                else
                {
                    Luis luis = new Luis();
                    luis.predict(e.PhraseResponse.Results[0].DisplayText, this);
                }
                
                ButtonStatus = "CLICK";
            }));
        }

        /// <summary>
        /// Writes the response result. 输出结果
        /// </summary>
        /// <param name="e">The <see cref="SpeechResponseEventArgs"/> instance containing the event data.</param>
        private void WriteResponseResult(SpeechResponseEventArgs e)
        {
            if (e.PhraseResponse.Results.Length == 0)
            {
                this.MessageShow = "No phrase response is available.";
            }
            else
            {
                this.SpeakPartialContent = e.PhraseResponse.Results[0].DisplayText;
            }
        }

        /// <summary>
        /// Called when a final response is received;
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechResponseEventArgs"/> instance containing the event data.</param>
        private void OnMicDictationResponseReceivedHandler(object sender, SpeechResponseEventArgs e)
        {
            // this.WriteLine("--- OnMicDictationResponseReceivedHandler ---");
            if (e.PhraseResponse.RecognitionStatus == RecognitionStatus.EndOfDictation ||
                e.PhraseResponse.RecognitionStatus == RecognitionStatus.DictationEndSilenceTimeout)
            {
                Dispatcher.Invoke(
                    (Action)(() =>
                    {
                        // we got the final result, so it we can end the mic reco.  No need to do this
                        // for dataReco, since we already called endAudio() on it as soon as we were done
                        // sending all the data.
                        this.micClient.EndMicAndRecognition();

                        this._startbutton.IsEnabled = true;
                        ButtonStatus = "CLICK";
                        //ResumeStyle();
                    }));

            }
            this.WriteResponseResult(e);
        }

        /// <summary>
        /// 作为动作结束的标志
        /// </summary>
        private bool done;


        /// <summary>
        /// Called when a final response is received and its intent is parsed
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechIntentEventArgs"/> instance containing the event data.</param>
        private void OnIntentHandler(object sender, SpeechIntentEventArgs e)
        {
            // 解析识别结果
            if (!done) return;
            done = false;
            Parser parser = new Parser();
            Parser.ParseResult parseResult = parser.getParseResult(e.Payload);
            // 识别结果传入Caller
            try
            {
                //Caller.CallFunc(parseResult);
            }
            catch (Exception error)
            {
                MessageBox.Show("Caller: " + error.Message);
            }
            finally
            {
                done = true;
            }
        }

        /// <summary>
        /// Called when a partial response is received.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="PartialSpeechResponseEventArgs"/> instance containing the event data.</param>
        private void OnPartialResponseReceivedHandler(object sender, PartialSpeechResponseEventArgs e)
        {
            // this.WriteLine("--- Partial result received by OnPartialResponseReceivedHandler() ---");
            // refresh speak information
            this.SpeakPartialContent = e.PartialResult;
        }

        /// <summary>
        /// Called when an error is received.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechErrorEventArgs"/> instance containing the event data.</param>
        private void OnConversationErrorHandler(object sender, SpeechErrorEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                _startbutton.IsEnabled = true;
                ButtonStatus = "CLICK";
            });
            //this.SpeakPartialContent = e.SpeechErrorText;
            this.MessageShow = e.SpeechErrorText;
        }

        /// <summary>
        /// 麦克风状态改变时调用
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="MicrophoneEventArgs"/> instance containing the event data.</param>
        private void OnMicrophoneStatus(object sender, MicrophoneEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                if (e.Recording)
                {
                    this.SpeakPartialContent = Properties.Resources.tips;
                }
            });
        }

        /// <summary>
        /// Helper function for INotifyPropertyChanged interface 
        /// </summary>
        /// <typeparam name="T">Property type</typeparam>
        /// <param name="caller">Property name</param>
        private void OnPropertyChanged<T>([CallerMemberName]string caller = null)
        {
            var handler = this.PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(caller));
            }
        }

    }
}
