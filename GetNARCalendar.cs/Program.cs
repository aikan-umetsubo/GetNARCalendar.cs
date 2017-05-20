using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

using HtmlAgilityPack;
using Ical.Net;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;
using Ical.Net.Serialization.iCalendar.Serializers;

namespace GetNARCalendar
{
	class GetNARCalendar
	{
		const string urlFormat
			= "http://www2.keiba.go.jp/KeibaWeb/MonthlyConveneInfo/MonthlyConveneInfoTop?k_year={0}&k_month={1}";

		static void Main(string[] args)
		{
			// 開催予定のリスト
			List<Schedule> scheduleList = new List<Schedule>();

			// 開催年を決定
			// 現在の年か、引数で渡された年か
			int year;
			if (args.Length == 0 || !int.TryParse(args[0], out year))
				year = DateTime.Now.Year;

			// 開催年の1月から12月までループ
			for (int month = 1; month <= 12; month++)
			{
				// 地方競馬のスケジュールのWebページを取得
				var request
                    = (HttpWebRequest)WebRequest.Create(string.Format(urlFormat, year, month));

				using (var response = (HttpWebResponse)request.GetResponse())
				using (var reader = new StreamReader(response.GetResponseStream()))
					// WebページをScheduleクラスのリストに加工して追加
					scheduleList.AddRange(ParseSchedulePage(reader.ReadToEnd()));
			}

			// Scheduleクラスのリストをicsファイルとして出力
			WriteAsIcs(scheduleList);
		}

		/// <summary>
		/// WebページをScheduleのリストに変換
		/// </summary>
		/// <param name="page">Webページ</param>
		/// <returns></returns>
		public static List<Schedule> ParseSchedulePage(string page)
		{
			List<Schedule> scheduleList = new List<Schedule>();

			var htmlDoc = new HtmlDocument();
			htmlDoc.LoadHtml(page);

			// HTMLから年と月を取得
			// （name属性がk_year、k_monthであるselect要素の子要素から、
			//   selected属性がtrueであるoption要素を取得）
			int year = htmlDoc
				.DocumentNode
				.SelectSingleNode(@"//select[@name=""k_year""]/option[@selected=""true""]")
				.GetAttributeValue("value", -1);

			int month = htmlDoc
				.DocumentNode
				.SelectSingleNode(@"//select[@name=""k_month""]/option[@selected=""true""]")
				.GetAttributeValue("value", -1);

			// HTMLからカレンダー部分を抜き出す
			// （class属性がdbtblであるtd要素の子要素から、
			// 　class属性がdbitemであるtd要素を子要素に持つtableを取得）
			var calendarRows = htmlDoc
				.DocumentNode
				.SelectNodes(@"//td[@class=""dbtbl""]/table[1]/tr[td[@class=""dbitem""]]");
			string previousType = string.Empty;
			foreach (HtmlNode calendarRow in calendarRows)
			{
				// 競馬場の所属、競馬場、1ヶ月の開催情報を取得
				//string courseClass = row.SelectSingleNode(@"td[@class=""dbtitle""]")
				//    ?.InnerText
				//    .Replace("\n", string.Empty);

				string courseName = calendarRow.SelectSingleNode(@"td[@class=""dbitem""]")
					?.InnerText
					.Replace("\n", string.Empty);

				var data = calendarRow.SelectNodes(@"td[@class=""dbdata""]");

				// rowspanが指定されて先頭セルが結合されている場合を考慮する。
				// （競馬場の所属が取得できない場合は、前の行のものを使う）
				//if (courseClass == null)
				//    courseClass = previousType;
				//else
				//    previousType = courseClass;

				// 1ヶ月の開催情報でループ
				for (int day = 1; day <= data.Count; day++)
				{
					// 開催の種別を判定
					RaceType raceType;
					switch (data[day - 1].InnerText.Replace("\n", string.Empty).Trim())
					{
						case "●":
							raceType = RaceType.Standard;
							break;
						case "☆":
							raceType = RaceType.Nighter;
							break;
						case "Ｄ":
							raceType = RaceType.Dart;
							break;
						case "△":
							raceType = RaceType.Substitute;
							break;
						default:
							raceType = RaceType.Closed;
							break;
					}

					// 開催予定をリストに追加
					scheduleList.Add(new Schedule(
						Racecourse.GetRaceCourseFromName(courseName),
						year,
						month,
						day,
						RaceType.Standard));
				}
			}

			return scheduleList;
		}

		/// <summary>
		/// ScheduleのリストをiCalendar形式で標準出力に書き出し
		/// </summary>
		/// <param name="scheduleList"></param>
		public static void WriteAsIcs(List<Schedule> scheduleList)
		{
			// ScheduleのリストをiCalendar形式に変換
			var calendar = new Calendar();
			foreach (Schedule schedule in scheduleList)
			{
				calendar.Events.Add(new Event
				{
					DtStart = new CalDateTime(schedule.Date),
					DtEnd = new CalDateTime(schedule.Date),
					Created = CalDateTime.Now,
					Description = string.Empty,
					LastModified = CalDateTime.Now,
					Status = EventStatus.Confirmed,
					Summary = schedule.Course.Japanese + "競馬",
					Transparency = TransparencyType.Opaque
				});
			}

			// 標準出力に出力
			var serializer = new CalendarSerializer(new SerializationContext());
			var serializedCalendar = serializer.SerializeToString(calendar);
			Console.OutputEncoding = Encoding.UTF8;
			Console.Out.Write(serializedCalendar);

			return;
		}
	}

	/// <summary>
	/// 一日の開催予定
	/// </summary>
	class Schedule
	{
		private Racecourse _race_course;
		private DateTime _date;
		private RaceType _race_type;

		public Racecourse Course { get { return _race_course; } }
		public DateTime Date { get { return _date; } }
		public RaceType Type { get { return _race_type; } }

		public Schedule(Racecourse course, DateTime date, RaceType type)
		{
			_race_course = course;
			_date = date;
			_race_type = type;
		}

		public Schedule(Racecourse course, int year, int month, int day, RaceType type)
		{
			_race_course = course;
			_date = new DateTime(year, month, day);
			_race_type = type;
		}
	}

	/// <summary>
	/// 競馬場
	/// </summary>
	class Racecourse
	{
		public Racecourse(Racecourse course) { _course = course._course; }

		private Racecourse(_race_course course) { _course = course; }

		public static readonly Racecourse Obihiro = new Racecourse(_race_course.Obihiro);
		public static readonly Racecourse Monbetsu = new Racecourse(_race_course.Monbetsu);
		public static readonly Racecourse Sapporo = new Racecourse(_race_course.Sapporo);
		public static readonly Racecourse Morioka = new Racecourse(_race_course.Morioka);
		public static readonly Racecourse Mizusawa = new Racecourse(_race_course.Mizusawa);
		public static readonly Racecourse Urawa = new Racecourse(_race_course.Urawa);
		public static readonly Racecourse Funabashi = new Racecourse(_race_course.Funabashi);
		public static readonly Racecourse Oi = new Racecourse(_race_course.Oi);
		public static readonly Racecourse Kawasaki = new Racecourse(_race_course.Kawasaki);
		public static readonly Racecourse Kanazawa = new Racecourse(_race_course.Kanazawa);
		public static readonly Racecourse Kasamatsu = new Racecourse(_race_course.Kasamatsu);
		public static readonly Racecourse Nagoya = new Racecourse(_race_course.Nagoya);
		public static readonly Racecourse Chukyo = new Racecourse(_race_course.Chukyo);
		public static readonly Racecourse Sonoda = new Racecourse(_race_course.Sonoda);
		public static readonly Racecourse Himeji = new Racecourse(_race_course.Himeji);
		public static readonly Racecourse Kochi = new Racecourse(_race_course.Kochi);
		public static readonly Racecourse Saga = new Racecourse(_race_course.Saga);

		private _race_course _course;
		public string Japanese { get { return _racecourses[(int)_course]; } }

		private enum _race_course
		{
			Obihiro = 0,
			Monbetsu = 1,
			Sapporo = 2,
			Morioka = 3,
			Mizusawa = 4,
			Urawa = 5,
			Funabashi = 6,
			Oi = 7,
			Kawasaki = 8,
			Kanazawa = 9,
			Kasamatsu = 10,
			Nagoya = 11,
			Chukyo = 12,
			Sonoda = 13,
			Himeji = 14,
			Kochi = 15,
			Saga = 16
		}

		private static string[] _racecourses =
		{
			"帯広",
			"門別",
			"札幌",
			"盛岡",
			"水沢",
			"浦和",
			"船橋",
			"大井",
			"川崎",
			"金沢",
			"笠松",
			"名古屋",
			"中京",
			"園田",
			"姫路",
			"高知",
			"佐賀"
		};

		public static Racecourse GetRaceCourseFromName(string name)
		{
			switch (name)
			{
				case "帯広":
				case "帯広ば":
					return Obihiro;
				case "門別":
					return Monbetsu;
				case "札幌":
					return Sapporo;
				case "盛岡":
					return Morioka;
				case "水沢":
					return Mizusawa;
				case "浦和":
					return Urawa;
				case "船橋":
					return Funabashi;
				case "大井":
					return Oi;
				case "川崎":
					return Kawasaki;
				case "金沢":
					return Kanazawa;
				case "笠松":
					return Kasamatsu;
				case "名古屋":
					return Nagoya;
				case "中京":
					return Chukyo;
				case "園田":
					return Sonoda;
				case "姫路":
					return Himeji;
				case "高知":
					return Kochi;
				case "佐賀":
					return Saga;
				default:
					return null;
			}
		}

		/// <summary>
		/// 競馬場の分類
		/// </summary>
		public class CourseClass
		{
			public CourseClass(CourseClass courseClass) { _class = courseClass._class; }

			private CourseClass(_race_course_class courseClass) { _class = courseClass; }

			public static readonly CourseClass Banei = new CourseClass(_race_course_class.Banei);
			public static readonly CourseClass Hokkaido = new CourseClass(_race_course_class.Hokkaido);
			public static readonly CourseClass Iwate = new CourseClass(_race_course_class.Iwate);
			public static readonly CourseClass MinamiKanto = new CourseClass(_race_course_class.MinamiKanto);
			public static readonly CourseClass Kanazawa = new CourseClass(_race_course_class.Kanazawa);
			public static readonly CourseClass Tokai = new CourseClass(_race_course_class.Tokai);
			public static readonly CourseClass Hyogo = new CourseClass(_race_course_class.Hyogo);
			public static readonly CourseClass Kochi = new CourseClass(_race_course_class.Kochi);
			public static readonly CourseClass Kyushu = new CourseClass(_race_course_class.Kyushu);

			private _race_course_class _class;

			public string Japanese { get { return _race_course_classes[(int)_class]; } }

			private enum _race_course_class
			{
				Banei = 0,
				Hokkaido = 1,
				Iwate = 2,
				MinamiKanto = 3,
				Kanazawa = 4,
				Tokai = 5,
				Hyogo = 6,
				Kochi = 7,
				Kyushu = 8
			}

			private static string[] _race_course_classes = {
				"ばんえい",
				"ホッカイドウ",
				"岩手",
				"南関東",
				"金沢",
				"東海",
				"兵庫",
				"高知",
				"九州"
			};
		}
	}

	/// <summary>
	/// 開催の種別
	/// </summary>
	class RaceType
	{
		public RaceType(RaceType type) { _type = type._type; }

		private RaceType(_race_type type) { _type = type; }

		/// <summary>開催の種別</summary>
		private _race_type _type;

		/// <summary>開催の種別（日本語）</summary>
		public string Japanese { get { return _race_types_jp[(int)_type]; } }

		#region 開催の種別（定数）

		/// <summary>通常開催</summary>
		public static readonly RaceType Standard = new RaceType(_race_type.Standard);

		/// <summary>ナイター開催</summary>
		public static readonly RaceType Nighter = new RaceType(_race_type.Nighter);

		/// <summary>ダート交流重賞競走</summary>
		public static readonly RaceType Dart = new RaceType(_race_type.Dart);

		/// <summary>別の日に代替開催</summary>
		public static readonly RaceType Substitute = new RaceType(_race_type.Substitute);

		/// <summary>開催なし</summary>
		public static readonly RaceType Closed = new RaceType(_race_type.Closed);

		#endregion

		/// <summary>開催の種別</summary>
		private enum _race_type
		{
			/// <summary>通常開催</summary>
			Standard = 0,
			/// <summary>ナイター競馬</summary>
			Nighter = 1,
			/// <summary>ダート交流重賞競走</summary>
			Dart = 2,
			/// <summary>別の日に代替開催</summary>
			Substitute = 3,
			/// <summary>開催なし</summary>
			Closed = 4
		}

		/// <summary>開催の種別（日本語）</summary>
		private static string[] _race_types_jp = {
			"通常開催",
			"ナイター競馬",
			"ダート交流重賞競走",
			"別の日に代替開催",
			"開催なし"
		};
	}
}
