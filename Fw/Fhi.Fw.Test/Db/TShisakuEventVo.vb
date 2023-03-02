Namespace Db
    ''' <summary>
    ''' 試作イベント情報
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class TShisakuEventVo
        '' 試作イベントコード
        Private _ShisakuEventCode As String
        '' 試作車系
        Private _ShisakuShakeiCode As String
        '' 試作開発符号
        Private _ShisakuKaihatsuFugo As String
        '' 車系/開発符号マスター表示順
        Private _HyojijunNo As Nullable(Of Int32)
        '' 試作イベントフェーズ
        Private _ShisakuEventPhase As String
        '' 試作イベントフェーズ名
        Private _ShisakuEventPhaseName As String
        '' ユニット区分
        Private _UnitKbn As String
        '' 試作イベント名称
        Private _ShisakuEventName As String
        '' 製作次期FROM
        Private _SeisakujikiFrom As Nullable(Of Int32)
        '' 製作次期TO
        Private _SeisakujikiTo As Nullable(Of Int32)
        '' 製作台数・完成車
        Private _SeisakudaisuKanseisya As Nullable(Of Int32)
        '' 製作台数・W/B
        Private _SeisakudaisuWb As Nullable(Of Int32)
        '' 製作台数・製作中止
        Private _SeisakudaisuChushi As Nullable(Of Int32)
        '' 製作一覧発行№
        Private _SeisakuichiranHakouNo As String
        '' 製作一覧発行№改訂
        Private _SeisakuichiranHakouNoKai As Nullable(Of Int32)
        '' 発注の有無
        Private _HachuUmu As String
        '' イベント情報登録担当所属
        Private _EventTourokuTantoKa As String
        '' イベント情報登録担当者
        Private _EventTourokuTanto As String
        '' 設計社員番号
        Private _SekkeiShainNo As String
        '' 設計展開日
        Private _SekkeiTenkaibi As Nullable(Of Int32)
        '' 訂正処置〆切日
        Private _KaiteiSyochiShimekiribi As Nullable(Of Int32)
        '' 〆切日
        Private _Shimekiribi As Nullable(Of Int32)
        '' 完了日
        Private _Kanryoubi As Nullable(Of Int32)
        '' 中止日
        Private _Chuushibi As Nullable(Of Int32)
        '' データ区分
        Private _DataKbn As String
        '' ステータス
        Private _Status As String
        '' 総ブロック数
        Private _TotalBlockQty As Nullable(Of Int32)
        '' 完了数
        Private _TotalCompQty As Nullable(Of Int32)
        '' 手配帳作成日
        Private _TehaichouSakuseibi As Nullable(Of Int32)
        '' 最終改訂抽出日
        Private _LastKaiteiChusyutubi As Nullable(Of Int32)
        '' 作成ユーザーID
        Private _CreatedUserId As String
        '' 作成日
        Private _CreatedDate As String
        '' 作成時
        Private _CreatedTime As String
        '' 更新ユーザーID
        Private _UpdatedUserId As String
        '' 更新日
        Private _UpdatedDate As String
        '' 更新時間
        Private _UpdatedTime As String

        ''' <summary>試作イベントコード</summary>
        ''' <value>試作イベントコード</value>
        ''' <returns>試作イベントコード</returns>
        Public Property ShisakuEventCode() As String
            Get
                Return _ShisakuEventCode
            End Get
            Set(ByVal value As String)
                _ShisakuEventCode = value
            End Set
        End Property

        ''' <summary>試作車系</summary>
        ''' <value>試作車系</value>
        ''' <returns>試作車系</returns>
        Public Property ShisakuShakeiCode() As String
            Get
                Return _ShisakuShakeiCode
            End Get
            Set(ByVal value As String)
                _ShisakuShakeiCode = value
            End Set
        End Property

        ''' <summary>試作開発符号</summary>
        ''' <value>試作開発符号</value>
        ''' <returns>試作開発符号</returns>
        Public Property ShisakuKaihatsuFugo() As String
            Get
                Return _ShisakuKaihatsuFugo
            End Get
            Set(ByVal value As String)
                _ShisakuKaihatsuFugo = value
            End Set
        End Property

        ''' <summary>車系/開発符号マスター表示順</summary>
        ''' <value>車系/開発符号マスター表示順</value>
        ''' <returns>車系/開発符号マスター表示順</returns>
        Public Property HyojijunNo() As Nullable(Of Int32)
            Get
                Return _HyojijunNo
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _HyojijunNo = value
            End Set
        End Property

        ''' <summary>試作イベントフェーズ</summary>
        ''' <value>試作イベントフェーズ</value>
        ''' <returns>試作イベントフェーズ</returns>
        Public Property ShisakuEventPhase() As String
            Get
                Return _ShisakuEventPhase
            End Get
            Set(ByVal value As String)
                _ShisakuEventPhase = value
            End Set
        End Property

        ''' <summary>試作イベントフェーズ名</summary>
        ''' <value>試作イベントフェーズ名</value>
        ''' <returns>試作イベントフェーズ名</returns>
        Public Property ShisakuEventPhaseName() As String
            Get
                Return _ShisakuEventPhaseName
            End Get
            Set(ByVal value As String)
                _ShisakuEventPhaseName = value
            End Set
        End Property

        ''' <summary>ユニット区分</summary>
        ''' <value>ユニット区分</value>
        ''' <returns>ユニット区分</returns>
        Public Property UnitKbn() As String
            Get
                Return _UnitKbn
            End Get
            Set(ByVal value As String)
                _UnitKbn = value
            End Set
        End Property

        ''' <summary>試作イベント名称</summary>
        ''' <value>試作イベント名称</value>
        ''' <returns>試作イベント名称</returns>
        Public Property ShisakuEventName() As String
            Get
                Return _ShisakuEventName
            End Get
            Set(ByVal value As String)
                _ShisakuEventName = value
            End Set
        End Property

        ''' <summary>製作次期FROM</summary>
        ''' <value>製作次期FROM</value>
        ''' <returns>製作次期FROM</returns>
        Public Property SeisakujikiFrom() As Nullable(Of Int32)
            Get
                Return _SeisakujikiFrom
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _SeisakujikiFrom = value
            End Set
        End Property

        ''' <summary>製作次期TO</summary>
        ''' <value>製作次期TO</value>
        ''' <returns>製作次期TO</returns>
        Public Property SeisakujikiTo() As Nullable(Of Int32)
            Get
                Return _SeisakujikiTo
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _SeisakujikiTo = value
            End Set
        End Property

        ''' <summary>製作台数・完成車</summary>
        ''' <value>製作台数・完成車</value>
        ''' <returns>製作台数・完成車</returns>
        Public Property SeisakudaisuKanseisya() As Nullable(Of Int32)
            Get
                Return _SeisakudaisuKanseisya
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _SeisakudaisuKanseisya = value
            End Set
        End Property

        ''' <summary>製作台数・W/B</summary>
        ''' <value>製作台数・W/B</value>
        ''' <returns>製作台数・W/B</returns>
        Public Property SeisakudaisuWb() As Nullable(Of Int32)
            Get
                Return _SeisakudaisuWb
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _SeisakudaisuWb = value
            End Set
        End Property

        ''' <summary>製作台数・製作中止</summary>
        ''' <value>製作台数・製作中止</value>
        ''' <returns>製作台数・製作中止</returns>
        Public Property SeisakudaisuChushi() As Nullable(Of Int32)
            Get
                Return _SeisakudaisuChushi
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _SeisakudaisuChushi = value
            End Set
        End Property

        ''' <summary>製作一覧発行№</summary>
        ''' <value>製作一覧発行№</value>
        ''' <returns>製作一覧発行№</returns>
        Public Property SeisakuichiranHakouNo() As String
            Get
                Return _SeisakuichiranHakouNo
            End Get
            Set(ByVal value As String)
                _SeisakuichiranHakouNo = value
            End Set
        End Property

        ''' <summary>製作一覧発行№改訂</summary>
        ''' <value>製作一覧発行№改訂</value>
        ''' <returns>製作一覧発行№改訂</returns>
        Public Property SeisakuichiranHakouNoKai() As Nullable(Of Int32)
            Get
                Return _SeisakuichiranHakouNoKai
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _SeisakuichiranHakouNoKai = value
            End Set
        End Property

        ''' <summary>発注の有無</summary>
        ''' <value>発注の有無</value>
        ''' <returns>発注の有無</returns>
        Public Property HachuUmu() As String
            Get
                Return _HachuUmu
            End Get
            Set(ByVal value As String)
                _HachuUmu = value
            End Set
        End Property

        ''' <summary>イベント情報登録担当所属</summary>
        ''' <value>イベント情報登録担当所属</value>
        ''' <returns>イベント情報登録担当所属</returns>
        Public Property EventTourokuTantoKa() As String
            Get
                Return _EventTourokuTantoKa
            End Get
            Set(ByVal value As String)
                _EventTourokuTantoKa = value
            End Set
        End Property

        ''' <summary>イベント情報登録担当者</summary>
        ''' <value>イベント情報登録担当者</value>
        ''' <returns>イベント情報登録担当者</returns>
        Public Property EventTourokuTanto() As String
            Get
                Return _EventTourokuTanto
            End Get
            Set(ByVal value As String)
                _EventTourokuTanto = value
            End Set
        End Property

        ''' <summary>設計社員番号</summary>
        ''' <value>設計社員番号</value>
        ''' <returns>設計社員番号</returns>
        Public Property SekkeiShainNo() As String
            Get
                Return _SekkeiShainNo
            End Get
            Set(ByVal value As String)
                _SekkeiShainNo = value
            End Set
        End Property

        ''' <summary>設計展開日</summary>
        ''' <value>設計展開日</value>
        ''' <returns>設計展開日</returns>
        Public Property SekkeiTenkaibi() As Nullable(Of Int32)
            Get
                Return _SekkeiTenkaibi
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _SekkeiTenkaibi = value
            End Set
        End Property

        ''' <summary>訂正処置〆切日</summary>
        ''' <value>訂正処置〆切日</value>
        ''' <returns>訂正処置〆切日</returns>
        Public Property KaiteiSyochiShimekiribi() As Nullable(Of Int32)
            Get
                Return _KaiteiSyochiShimekiribi
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _KaiteiSyochiShimekiribi = value
            End Set
        End Property

        ''' <summary>受付日</summary>
        ''' <value>受付日</value>
        ''' <returns>受付日</returns>
        Public Property Shimekiribi() As Nullable(Of Int32)
            Get
                Return _Shimekiribi
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _Shimekiribi = value
            End Set
        End Property

        ''' <summary>完了日</summary>
        ''' <value>完了日</value>
        ''' <returns>完了日</returns>
        Public Property Kanryoubi() As Nullable(Of Int32)
            Get
                Return _Kanryoubi
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _Kanryoubi = value
            End Set
        End Property

        ''' <summary>中止日</summary>
        ''' <value>中止日</value>
        ''' <returns>中止日</returns>
        Public Property Chuushibi() As Nullable(Of Int32)
            Get
                Return _Chuushibi
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _Chuushibi = value
            End Set
        End Property

        ''' <summary>データ区分</summary>
        ''' <value>データ区分</value>
        ''' <returns>データ区分</returns>
        Public Property DataKbn() As String
            Get
                Return _DataKbn
            End Get
            Set(ByVal value As String)
                _DataKbn = value
            End Set
        End Property

        ''' <summary>ステータス</summary>
        ''' <value>ステータス</value>
        ''' <returns>ステータス</returns>
        Public Property Status() As String
            Get
                Return _Status
            End Get
            Set(ByVal value As String)
                _Status = value
            End Set
        End Property

        ''' <summary>総ブロック数</summary>
        ''' <value>総ブロック数</value>
        ''' <returns>総ブロック数</returns>
        Public Property TotalBlockQty() As Nullable(Of Int32)
            Get
                Return _TotalBlockQty
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _TotalBlockQty = value
            End Set
        End Property

        ''' <summary>完了数</summary>
        ''' <value>完了数</value>
        ''' <returns>完了数</returns>
        Public Property TotalCompQty() As Nullable(Of Int32)
            Get
                Return _TotalCompQty
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _TotalCompQty = value
            End Set
        End Property

        ''' <summary>手配帳作成日</summary>
        ''' <value>手配帳作成日</value>
        ''' <returns>手配帳作成日</returns>
        Public Property TehaichouSakuseibi() As Nullable(Of Int32)
            Get
                Return _TehaichouSakuseibi
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _TehaichouSakuseibi = value
            End Set
        End Property

        ''' <summary>最終改訂抽出日</summary>
        ''' <value>最終改訂抽出日</value>
        ''' <returns>最終改訂抽出日</returns>
        Public Property LastKaiteiChusyutubi() As Nullable(Of Int32)
            Get
                Return _LastKaiteiChusyutubi
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _LastKaiteiChusyutubi = value
            End Set
        End Property

        ''' <summary>作成ユーザーID</summary>
        ''' <value>作成ユーザーID</value>
        ''' <returns>作成ユーザーID</returns>
        Public Property CreatedUserId() As String
            Get
                Return _CreatedUserId
            End Get
            Set(ByVal value As String)
                _CreatedUserId = value
            End Set
        End Property

        ''' <summary>作成日</summary>
        ''' <value>作成日</value>
        ''' <returns>作成日</returns>
        Public Property CreatedDate() As String
            Get
                Return _CreatedDate
            End Get
            Set(ByVal value As String)
                _CreatedDate = value
            End Set
        End Property

        ''' <summary>作成時</summary>
        ''' <value>作成時</value>
        ''' <returns>作成時</returns>
        Public Property CreatedTime() As String
            Get
                Return _CreatedTime
            End Get
            Set(ByVal value As String)
                _CreatedTime = value
            End Set
        End Property

        ''' <summary>更新ユーザーID</summary>
        ''' <value>更新ユーザーID</value>
        ''' <returns>更新ユーザーID</returns>
        Public Property UpdatedUserId() As String
            Get
                Return _UpdatedUserId
            End Get
            Set(ByVal value As String)
                _UpdatedUserId = value
            End Set
        End Property

        ''' <summary>更新日</summary>
        ''' <value>更新日</value>
        ''' <returns>更新日</returns>
        Public Property UpdatedDate() As String
            Get
                Return _UpdatedDate
            End Get
            Set(ByVal value As String)
                _UpdatedDate = value
            End Set
        End Property

        ''' <summary>更新時間</summary>
        ''' <value>更新時間</value>
        ''' <returns>更新時間</returns>
        Public Property UpdatedTime() As String
            Get
                Return _UpdatedTime
            End Get
            Set(ByVal value As String)
                _UpdatedTime = value
            End Set
        End Property
    End Class
End Namespace