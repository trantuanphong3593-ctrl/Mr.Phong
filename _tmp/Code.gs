/**
 * ============================================================================
 *  WORD FORM 9 BY NMC — MÁY CHỦ GOOGLE APPS SCRIPT (Code.gs)  •  Phiên bản 5 (4 dạng bài)
 * ============================================================================
 *  Script này gắn với Google Sheet của giáo viên (Tiện ích mở rộng > Apps Script).
 *
 *  GIÁO VIÊN CHỈ CẦN LÀM VIỆC TRONG GOOGLE SHEET:
 *   - DanhSach   : danh sách học sinh (lớp, họ tên). Có menu nhập nhanh bằng cách dán.
 *   - CauHoi     : ngân hàng câu hỏi (đề bài, từ gốc, đáp án, giải thích, dịch nghĩa...).
 *                  Câu nào CHƯA CÓ ĐÁP ÁN thì học sinh được TỰ CHẤM; thêm đáp án sau
 *                  rồi chọn menu "Chấm lại" là máy chấm lại cho tất cả học sinh.
 *   - GiaDinhTu  : bảng gia đình từ (dùng cho gợi ý và Sổ tay).
 *   - TuDien     : các mục từ điển dùng cho dạng bài Word entry.
 *   Cột "Dạng bài" của CauHoi: Điền từ | Trắc nghiệm | Viết lại câu | Word entry (để trống = Điền từ).
 *   - TienDo     : kết quả mỗi học sinh một dòng + điểm thang 10 (tự động).
 *   - NhatKy / TongQuan / ThongKeCau : nhật ký, thống kê lớp, thống kê câu (tự động).
 *
 *  Đáp án KHÔNG nằm trong mã trang của học sinh: máy chủ chấm rồi mới trả đáp án.
 *  Cách triển khai: xem file HUONG_DAN.md
 * ============================================================================
 */

var TZ = 'Asia/Ho_Chi_Minh';
var MAX_Q = 2000; // số câu tối đa (STT từ 1 đến 2000)

var SH_ROSTER = 'DanhSach';
var SH_QUESTIONS = 'CauHoi';
var SH_FAMILY = 'GiaDinhTu';
var SH_ENTRY = 'TuDien';
var Q_COLS = 18;    // số cột của sheet CauHoi (A..R); P = Unit, Q = Chủ đề, R = Áp dụng cho (khối/lớp)
var SH_PROGRESS = 'TienDo';
var SH_LOG = 'NhatKy';
var SH_WATCH = 'GiamSat';          // v8: nhật ký rời trang (chống gian lận)
var AC_PROP = 'AC_CFG';           // v8: cài đặt chống gian lận trong Script Properties
var SH_OVERVIEW = 'TongQuan';
var SH_QSTATS = 'ThongKeCau';

// Bố cục sheet TienDo (1-based):
// A Lớp | B Họ tên | C Đã làm | D Đúng | E Sai | F Chờ chấm | G Tiến độ | H % đúng |
// I Điểm/10 | J Tổng lượt | K Lần cuối | L..CI = Câu 1..100 | CJ (cột 112) = JSON bài làm |
// CK.. (cột 113..2012) = Câu 101..2000. Cột JSON giữ nguyên vị trí từ các bản trước nên dữ liệu cũ vẫn dùng được.
var COL_Q0 = 12;
var COL_TEXTS = 112;
var Q_BLOCK1 = 100;
var PROG_COLS = COL_TEXTS + (MAX_Q - Q_BLOCK1);   // 2012

/** Số cột (1-based) của câu có STT = id trong sheet TienDo. */
function qCol_(id) { return id <= Q_BLOCK1 ? COL_Q0 - 1 + id : COL_TEXTS + (id - Q_BLOCK1); }

/** Gợi ý Unit (sách Tiếng Anh 9 Global Success) cho ô chọn nhanh; giáo viên vẫn có thể gõ Unit/chủ đề tùy ý. */
var UNIT_SUGGEST = [
  'Unit 1: Local Environment', 'Unit 2: City Life', 'Unit 3: Teen Stress and Pressure', 'Unit 4: Life in the Past',
  'Unit 5: Wonders of Viet Nam', 'Unit 6: Viet Nam: Then and Now', 'Unit 7: Recipes and Eating Habits', 'Unit 8: Tourism',
  'Unit 9: English in the World', 'Unit 10: Space Travel', 'Unit 11: Changing Roles in Society', 'Unit 12: My Future Career',
  'Ôn tổng hợp'
];

var ST_OK = 'Đ';        // máy chấm: đúng
var ST_BAD = 'S';       // máy chấm: sai
var ST_REV = 'X';       // học sinh xem đáp án
var ST_PEND = '?';      // chưa có đáp án, chờ giáo viên chấm
var ST_SELF_OK = 'T+';  // học sinh tự chấm: đúng
var ST_SELF_BAD = 'T-'; // học sinh tự chấm: sai

// Câu mẫu (bản 6): 40 câu điền từ của file gốc, 8 cột đầu theo sheet CauHoi
var SAMPLE_QUESTIONS = [
 [
  1,
  "The traditional craft village has become one of the most popular tourist __________ in the region.",
  "ATTRACT",
  "attractions",
  "",
  "Sau cấu trúc 'one of the most popular tourist...' cần một danh từ đếm được số nhiều. 'Tourist attractions' nghĩa là các điểm thu hút khách du lịch.",
  "Làng nghề truyền thống đã trở thành một trong những điểm thu hút khách du lịch phổ biến nhất trong vùng.",
  "ATTRACT"
 ],
 [
  2,
  "Local volunteers worked __________ and completed all the preparations on time.",
  "EFFECT",
  "effectively",
  "",
  "Sau động từ thường 'worked' cần một trạng từ chỉ cách thức để bổ nghĩa cho hành động làm việc như thế nào.",
  "Các tình nguyện viên địa phương đã làm việc một cách hiệu quả và hoàn thành mọi công tác chuẩn bị đúng hạn.",
  "EFFECT"
 ],
 [
  3,
  "The city government plans to __________ public transport services in crowded residential areas.",
  "DEVELOPMENT",
  "develop",
  "",
  "Sau cấu trúc 'plans to + V (nguyên mẫu)' cần một động từ nguyên mẫu. Từ gốc 'development' (n) chuyển về động từ là 'develop'.",
  "Chính quyền thành phố có kế hoạch phát triển các dịch vụ giao thông công cộng tại những khu dân cư đông đúc.",
  "DEVELOP"
 ],
 [
  4,
  "Many teenagers feel __________ when they have too many examinations in the same week.",
  "STRESS",
  "stressed",
  "",
  "Sau động từ liên kết 'feel' cần một tính từ. Chủ ngữ là người ('Many teenagers') chịu tác động tâm lý nên dùng tính từ đuôi '-ed' (stressed: cảm thấy căng thẳng).",
  "Nhiều thanh thiếu niên cảm thấy căng thẳng khi có quá nhiều kỳ thi trong cùng một tuần.",
  "STRESS"
 ],
 [
  5,
  "The lack of nearby medical services causes considerable __________ for residents in remote villages.",
  "CONVENIENT",
  "inconvenience",
  "",
  "Sau tính từ 'considerable' (đáng kể) cần một danh từ. Dựa vào ngữ cảnh 'thiếu dịch vụ y tế gần nhà' nên phải mang nghĩa tiêu cực là 'inconvenience' (sự bất tiện).",
  "Việc thiếu các dịch vụ y tế gần đó gây ra sự bất tiện đáng kể cho người dân ở các ngôi làng hẻo lánh.",
  "CONVENIENT"
 ],
 [
  6,
  "City authorities should __________ residents to use public transport more often.",
  "COURAGE",
  "encourage",
  "",
  "Sau động từ khuyết thiếu 'should' cần một động từ nguyên mẫu. Cấu trúc 'encourage somebody to V': khuyến khích ai làm gì.",
  "Chính quyền thành phố nên khuyến khích người dân sử dụng phương tiện giao thông công cộng thường xuyên hơn.",
  "COURAGE"
 ],
 [
  7,
  "The old bus service was __________ because buses often arrived late or failed to arrive at all.",
  "RELY",
  "unreliable",
  "",
  "Sau động từ 'to be' (was) cần một tính từ. Mệnh đề sau 'because' giải thích xe buýt thường đến trễ hoặc bỏ chuyến nên cần tính từ mang nghĩa phủ định là 'unreliable' (không đáng tin cậy).",
  "Dịch vụ xe buýt cũ không đáng tin cậy vì xe buýt thường đến trễ hoặc thậm chí không đến.",
  "RELY"
 ],
 [
  8,
  "Students must follow basic __________ safety rules when using computers and other devices.",
  "ELECTRIC",
  "electrical",
  "",
  "Trước cụm danh từ 'safety rules' cần một tính từ. Khi nói về quy tắc an toàn liên quan đến điện nói chung, ta dùng cụm cố định 'electrical safety rules' (quy tắc an toàn điện).",
  "Học sinh phải tuân theo các quy tắc an toàn điện cơ bản khi sử dụng máy tính và các thiết bị khác.",
  "ELECTRIC"
 ],
 [
  9,
  "Students should __________ their schoolwork with leisure activities.",
  "BALANCED",
  "balance",
  "",
  "Sau động từ khuyết thiếu 'should' cần một động từ nguyên mẫu. Cấu trúc 'balance A with B': cân bằng giữa A và B.",
  "Học sinh nên cân bằng việc học ở trường với các hoạt động giải trí.",
  "BALANCE"
 ],
 [
  10,
  "The local council responded __________ to residents’ complaints by arranging an immediate clean-up.",
  "APPROPRIATE",
  "appropriately",
  "",
  "Sau nội động từ 'responded' cần một trạng từ bổ nghĩa cho hành động phản hồi như thế nào ('appropriately': một cách phù hợp, thỏa đáng).",
  "Hội đồng địa phương đã phản hồi một cách thỏa đáng trước phàn nàn của người dân bằng việc thu xếp dọn dẹp ngay lập tức.",
  "APPROPRIATE"
 ],
 [
  11,
  "The city government hopes to __________ traffic congestion by improving public transport.",
  "REDUCTION",
  "reduce",
  "",
  "Sau cấu trúc 'hopes to + V (nguyên mẫu)' cần một động từ nguyên mẫu. Danh từ 'reduction' chuyển về động từ là 'reduce' (cắt giảm, làm giảm).",
  "Chính quyền thành phố hy vọng sẽ giảm ùn tắc giao thông bằng cách cải thiện giao thông công cộng.",
  "REDUCE"
 ],
 [
  12,
  "Many teenagers experience a great deal of __________ before important examinations.",
  "ANXIOUS",
  "anxiety",
  "",
  "Sau cụm lượng từ 'a great deal of' (nhiều - đi với danh từ không đếm được) cần một danh từ. Tính từ 'anxious' chuyển sang danh từ là 'anxiety' (sự lo lắng).",
  "Nhiều thanh thiếu niên trải qua rất nhiều sự lo lắng trước những kỳ thi quan trọng.",
  "ANXIOUS"
 ],
 [
  13,
  "Residents should take __________ for keeping their neighbourhood clean.",
  "RESPONSIBLE",
  "responsibility",
  "",
  "Cụm cố định 'take responsibility for + V-ing/N' nghĩa là chịu trách nhiệm về việc gì.",
  "Người dân nên có trách nhiệm trong việc giữ gìn khu phố của mình sạch sẽ.",
  "RESPONSIBLE"
 ],
 [
  14,
  "The school encourages students to __________ in community projects during the summer holidays.",
  "PARTICIPATION",
  "participate",
  "",
  "Sau cấu trúc 'encourages somebody to + V (nguyên mẫu)' và đi kèm giới từ 'in', ta cần động từ nguyên mẫu 'participate' (participate in: tham gia vào).",
  "Nhà trường khuyến khích học sinh tham gia vào các dự án cộng đồng trong kỳ nghỉ hè.",
  "PARTICIPATE"
 ],
 [
  15,
  "The main street became extremely __________ during the annual festival.",
  "CONGEST",
  "congested",
  "",
  "Sau động từ liên kết 'became' và trạng từ chỉ mức độ 'extremely' cần một tính từ. 'Congested' (adj) nghĩa là bị tắc nghẽn, đông nghịt.",
  "Con phố chính trở nên cực kỳ tắc nghẽn trong suốt lễ hội thường niên.",
  "CONGEST"
 ],
 [
  16,
  "Local artists decorated the community centre __________ with colourful flowers and handmade lanterns.",
  "ATTRACT",
  "attractively",
  "",
  "Câu đã có đủ động từ 'decorated' và tân ngữ 'the community centre', cần một trạng từ ('attractively') để bổ nghĩa cho hành động trang trí đẹp mắt như thế nào.",
  "Các nghệ sĩ địa phương đã trang trí trung tâm cộng đồng một cách bắt mắt bằng những bông hoa rực rỡ và đèn lồng thủ công.",
  "ATTRACT"
 ],
 [
  17,
  "__________ important tasks helps teenagers manage their study time better.",
  "PRIORITY",
  "Prioritizing",
  "Prioritising|prioritizing|prioritising",
  "Đóng vai trò làm chủ ngữ đứng đầu câu (trước tân ngữ 'important tasks' và động từ chia số ít 'helps') cần một Danh động từ (V-ing). Động từ của 'priority' là 'prioritize/prioritise' → V-ing: 'Prioritizing / Prioritising'.",
  "Việc ưu tiên các nhiệm vụ quan trọng giúp thanh thiếu niên quản lý thời gian học tập tốt hơn.",
  "PRIORITY"
 ],
 [
  18,
  "Air __________ has become a serious problem in the city.",
  "POLLUTE",
  "pollution",
  "",
  "Kết hợp với danh từ 'Air' để tạo thành cụm danh từ làm chủ ngữ ('Air pollution': sự ô nhiễm không khí) đứng trước động từ 'has become'.",
  "Ô nhiễm không khí đã trở thành một vấn đề nghiêm trọng trong thành phố.",
  "POLLUTE"
 ],
 [
  19,
  "The school plans to __________ its sports facilities next year.",
  "IMPROVEMENT",
  "improve",
  "",
  "Sau cấu trúc 'plans to + V (nguyên mẫu)' cần một động từ nguyên mẫu. Danh từ 'improvement' chuyển về động từ gốc là 'improve'.",
  "Nhà trường có kế hoạch nâng cấp cơ sở vật chất thể thao vào năm tới.",
  "IMPROVE"
 ],
 [
  20,
  "The teenager __________ ignored safety warnings and rode through a red light.",
  "RESPONSIBLE",
  "irresponsibly",
  "",
  "Đứng trước động từ thường 'ignored' cần một trạng từ bổ nghĩa cho hành động. Do ngữ cảnh 'phớt lờ cảnh báo an toàn và vượt đèn đỏ' nên cần trạng từ mang nghĩa phủ định 'irresponsibly' (một cách vô trách nhiệm).",
  "Thiếu niên đó đã phớt lờ các cảnh báo an toàn một cách vô trách nhiệm và vượt đèn đỏ.",
  "RESPONSIBLE"
 ],
 [
  21,
  "The police are investigating several __________ activities near the railway station.",
  "CRIME",
  "criminal",
  "",
  "Trước danh từ 'activities' cần một tính từ bổ nghĩa. 'Criminal activities' nghĩa là các hoạt động tội phạm / phạm pháp.",
  "Cảnh sát đang điều tra một số hoạt động phạm pháp gần nhà ga xe lửa.",
  "CRIME"
 ],
 [
  22,
  "Parents should __________ their children to take regular breaks while studying.",
  "ENCOURAGEMENT",
  "encourage",
  "",
  "Sau động từ khuyết thiếu 'should' cần một động từ nguyên mẫu. Cấu trúc 'encourage somebody to V': khuyến khích ai làm gì.",
  "Cha mẹ nên khuyến khích con cái nghỉ giải lao đều đặn trong khi học.",
  "COURAGE"
 ],
 [
  23,
  "__________, the new community library was named after a famous local writer.",
  "APPROPRIATE",
  "Appropriately",
  "",
  "Đứng đầu câu và ngăn cách với mệnh đề chính bằng dấu phẩy cần một trạng từ bổ nghĩa cho cả câu ('Appropriately': Thật phù hợp / Đáng lý thay).",
  "Thật phù hợp khi thư viện cộng đồng mới được đặt theo tên của một nhà văn nổi tiếng ở địa phương.",
  "APPROPRIATE"
 ],
 [
  24,
  "A __________ daily timetable gives teenagers enough time for study, exercise, and sleep.",
  "BALANCE",
  "balanced",
  "",
  "Đứng sau mạo từ 'A' và trước cụm danh từ 'daily timetable' cần một tính từ. 'A balanced daily timetable' là một thời gian biểu hằng ngày cân đối/hợp lý.",
  "Một thời gian biểu hằng ngày cân đối giúp thanh thiếu niên có đủ thời gian để học tập, tập thể dục và ngủ nghỉ.",
  "BALANCE"
 ],
 [
  25,
  "Community members should __________ local traditions by teaching traditional crafts to children.",
  "PRESERVATION",
  "preserve",
  "",
  "Sau động từ khuyết thiếu 'should' cần một động từ nguyên mẫu. Danh từ 'preservation' chuyển về động từ gốc là 'preserve' (bảo tồn, giữ gìn).",
  "Các thành viên trong cộng đồng nên bảo tồn truyền thống địa phương bằng cách dạy nghề thủ công truyền thống cho trẻ em.",
  "PRESERVE"
 ],
 [
  26,
  "Constant traffic noise is a major source of __________ for students living downtown.",
  "DISTRACT",
  "distraction",
  "",
  "Sau giới từ 'of' trong cụm 'a major source of...' cần một danh từ. Động từ 'distract' chuyển thành danh từ 'distraction' (sự xao nhãng).",
  "Tiếng ồn giao thông liên tục là một nguyên nhân lớn gây xao nhãng cho học sinh sống ở khu trung tâm.",
  "DISTRACT"
 ],
 [
  27,
  "The school decided to __________ students with modern computers for digital learning.",
  "EQUIPMENT",
  "equip",
  "",
  "Sau cấu trúc 'decided to + V (nguyên mẫu)' cần một động từ nguyên mẫu. Cấu trúc 'equip somebody with something': trang bị cho ai cái gì.",
  "Nhà trường quyết định trang bị cho học sinh những chiếc máy tính hiện đại để phục vụ việc học tập số.",
  "EQUIP"
 ],
 [
  28,
  "Online health advice is not always __________ because some websites provide incorrect information.",
  "RELY",
  "reliable",
  "",
  "Sau động từ 'to be' (is not always...) cần một tính từ. Vì phía trước đã có từ phủ định 'not', ta điền tính từ khẳng định 'reliable' (not always reliable: không phải lúc nào cũng đáng tin cậy).",
  "Lời khuyên sức khỏe trực tuyến không phải lúc nào cũng đáng tin cậy vì một số trang web cung cấp thông tin không chính xác.",
  "RELY"
 ],
 [
  29,
  "Students worked __________ and completed more tasks after learning time-management skills.",
  "PRODUCE",
  "productively",
  "",
  "Sau động từ thường 'worked' cần một trạng từ bổ nghĩa cho hành động làm việc ('productively': một cách năng suất, hiệu quả).",
  "Học sinh đã làm việc hiệu quả hơn và hoàn thành được nhiều nhiệm vụ hơn sau khi học kỹ năng quản lý thời gian.",
  "PRODUCE"
 ],
 [
  30,
  "The local government hopes to __________ the old village into a popular cultural destination.",
  "DEVELOPMENT",
  "develop",
  "",
  "Sau cấu trúc 'hopes to + V (nguyên mẫu)' cần một động từ nguyên mẫu. Cấu trúc 'develop something into something': phát triển cái gì thành cái gì.",
  "Chính quyền địa phương hy vọng sẽ phát triển ngôi làng cổ thành một điểm đến văn hóa nổi tiếng.",
  "DEVELOP"
 ],
 [
  31,
  "Too much homework can make teenagers __________, especially during examination periods.",
  "STRESS",
  "stressed",
  "",
  "Cấu trúc 'make + O + Adj' (làm cho ai trở nên như thế nào). Tân ngữ là người ('teenagers') nên dùng tính từ đuôi '-ed' ('stressed': cảm thấy căng thẳng).",
  "Quá nhiều bài tập về nhà có thể khiến thanh thiếu niên bị căng thẳng, đặc biệt là trong các giai đoạn thi cử.",
  "STRESS"
 ],
 [
  32,
  "The community centre hired an experienced __________ to repair the damaged lights.",
  "ELECTRIC",
  "electrician",
  "",
  "Sau mạo từ 'an' và tính từ 'experienced' (giàu kinh nghiệm), đồng thời chỉ người thực hiện hành động 'repair the damaged lights' (sửa đèn hỏng), cần danh từ chỉ người số ít: 'electrician' (thợ điện).",
  "Trung tâm cộng đồng đã thuê một thợ điện giàu kinh nghiệm để sửa chữa những chiếc đèn bị hỏng.",
  "ELECTRIC"
 ],
 [
  33,
  "__________ unnecessary distractions helps students concentrate on their schoolwork.",
  "AVOIDANCE",
  "Avoiding",
  "",
  "Đứng đầu câu làm chủ ngữ và theo sau trực tiếp là tân ngữ 'unnecessary distractions', ta cần một Danh động từ (V-ing) là 'Avoiding' (Việc tránh...). Danh từ 'Avoidance' không đi trực tiếp với tân ngữ mà không có giới từ 'of'.",
  "Việc tránh những yếu tố gây xao nhãng không cần thiết giúp học sinh tập trung vào việc học.",
  "AVOID"
 ],
 [
  34,
  "The city council plans to __________ several historic buildings before they become seriously damaged.",
  "PRESERVATION",
  "preserve",
  "",
  "Sau cấu trúc 'plans to + V (nguyên mẫu)' cần một động từ nguyên mẫu ('preserve': bảo tồn).",
  "Hội đồng thành phố có kế hoạch bảo tồn một số tòa nhà lịch sử trước khi chúng bị hư hại nghiêm trọng.",
  "PRESERVE"
 ],
 [
  35,
  "Residents complained about the __________ caused by frequent bus cancellations.",
  "CONVENIENT",
  "inconvenience",
  "",
  "Sau mạo từ 'the' cần một danh từ. Do ngữ cảnh 'phàn nàn về việc hủy chuyến xe buýt thường xuyên' nên cần danh từ mang nghĩa tiêu cực 'inconvenience' (sự bất tiện).",
  "Người dân phàn nàn về sự bất tiện do việc hủy chuyến xe buýt thường xuyên gây ra.",
  "CONVENIENT"
 ],
 [
  36,
  "Teachers should speak __________ to students, using language suitable for their age.",
  "APPROPRIATE",
  "appropriately",
  "",
  "Sau động từ thường 'speak' cần một trạng từ chỉ cách thức để bổ nghĩa cho hành động nói chuyện ('appropriately': một cách phù hợp, đúng mực).",
  "Giáo viên nên nói chuyện một cách phù hợp với học sinh, sử dụng ngôn từ vừa với lứa tuổi của các em.",
  "APPROPRIATE"
 ],
 [
  37,
  "The new road was built to reduce traffic __________ in the city centre.",
  "CONGEST",
  "congestion",
  "",
  "Kết hợp với danh từ 'traffic' tạo thành cụm danh từ quen thuộc 'traffic congestion' (sự ùn tắc giao thông) làm tân ngữ cho động từ 'reduce'.",
  "Con đường mới được xây dựng nhằm giảm thiểu ùn tắc giao thông ở trung tâm thành phố.",
  "CONGEST"
 ],
 [
  38,
  "Regular outdoor activities can __________ students’ physical health.",
  "IMPROVEMENT",
  "improve",
  "",
  "Sau động từ khuyết thiếu 'can' cần một động từ nguyên mẫu ('improve': cải thiện).",
  "Các hoạt động ngoài trời thường xuyên có thể cải thiện sức khỏe thể chất của học sinh.",
  "IMPROVE"
 ],
 [
  39,
  "With the mayor’s __________, more residents joined the environmental protection campaign.",
  "COURAGE",
  "encouragement",
  "",
  "Sau sở hữu cách 'the mayor's' cần một danh từ. Dựa vào ngữ cảnh 'nhiều người dân tham gia chiến dịch bảo vệ môi trường hơn', cần danh từ 'encouragement' (sự khuyến khích, khích lệ).",
  "Với sự khích lệ của thị trưởng, nhiều người dân đã tham gia chiến dịch bảo vệ môi trường hơn.",
  "COURAGE"
 ],
 [
  40,
  "The local government will __________ community centres with modern computers.",
  "EQUIPMENT",
  "equip",
  "",
  "Sau động từ khuyết thiếu 'will' cần một động từ nguyên mẫu. Cấu trúc 'equip A with B': trang bị cho A cái gì.",
  "Chính quyền địa phương sẽ trang bị máy tính hiện đại cho các trung tâm cộng đồng.",
  "EQUIP"
 ]
];

// Unit / Chủ đề gợi ý cho 40 câu trên (cùng thứ tự). Giáo viên có thể sửa trong sheet CauHoi (cột P, Q).
var SAMPLE_META_OLD = [["Unit 1: Local Environment", "Danh từ"], ["Unit 1: Local Environment", "Trạng từ"], ["Unit 2: City Life", "Động từ"], ["Unit 3: Teen Stress and Pressure", "Tính từ"], ["Unit 2: City Life", "Danh từ"], ["Unit 2: City Life", "Động từ"], ["Unit 2: City Life", "Tính từ"], ["Ôn tổng hợp", "Tính từ"], ["Unit 3: Teen Stress and Pressure", "Động từ"], ["Unit 2: City Life", "Trạng từ"], ["Unit 2: City Life", "Động từ"], ["Unit 3: Teen Stress and Pressure", "Danh từ"], ["Unit 1: Local Environment", "Danh từ"], ["Unit 1: Local Environment", "Động từ"], ["Unit 2: City Life", "Tính từ"], ["Unit 1: Local Environment", "Trạng từ"], ["Unit 3: Teen Stress and Pressure", "Danh động từ (V-ing)"], ["Unit 2: City Life", "Danh từ"], ["Ôn tổng hợp", "Động từ"], ["Ôn tổng hợp", "Trạng từ"], ["Unit 2: City Life", "Tính từ"], ["Unit 3: Teen Stress and Pressure", "Động từ"], ["Unit 1: Local Environment", "Trạng từ"], ["Unit 3: Teen Stress and Pressure", "Tính từ"], ["Unit 1: Local Environment", "Động từ"], ["Unit 3: Teen Stress and Pressure", "Danh từ"], ["Ôn tổng hợp", "Động từ"], ["Ôn tổng hợp", "Tính từ"], ["Unit 3: Teen Stress and Pressure", "Trạng từ"], ["Unit 1: Local Environment", "Động từ"], ["Unit 3: Teen Stress and Pressure", "Tính từ"], ["Unit 1: Local Environment", "Danh từ"], ["Unit 3: Teen Stress and Pressure", "Danh động từ (V-ing)"], ["Unit 2: City Life", "Động từ"], ["Unit 2: City Life", "Danh từ"], ["Ôn tổng hợp", "Trạng từ"], ["Unit 2: City Life", "Danh từ"], ["Ôn tổng hợp", "Động từ"], ["Unit 1: Local Environment", "Danh từ"], ["Unit 1: Local Environment", "Động từ"]];

var SAMPLE_FAMILIES = [
 [
  "ATTRACT",
  "Attract (v): thu hút, hấp dẫn → attraction (n): sự thu hút; điểm thu hút → attractions (n): các điểm thu hút → attractive (adj): hấp dẫn, thu hút → attractively (adv): một cách bắt mắt, hấp dẫn."
 ],
 [
  "EFFECT",
  "Effect (n): hiệu quả, tác động → effective (adj): có hiệu quả → ineffective (adj): không hiệu quả → effectively (adv): một cách hiệu quả."
 ],
 [
  "DEVELOP",
  "Develop (v): phát triển, mở rộng → development (n): sự phát triển → developing (adj): đang phát triển → developed (adj): đã phát triển."
 ],
 [
  "STRESS",
  "Stress (n, v): sự căng thẳng; gây căng thẳng → stressed (adj): cảm thấy căng thẳng → stressful (adj): gây căng thẳng."
 ],
 [
  "CONVENIENT",
  "Convenient (adj): thuận tiện, tiện lợi → inconvenient (adj): bất tiện → convenience (n): sự tiện lợi → inconvenience (n): sự bất tiện → conveniently (adv): một cách thuận tiện."
 ],
 [
  "COURAGE",
  "Courage (n): lòng can đảm → encourage (v): khuyến khích, động viên → encouragement (n): sự khuyến khích, khích lệ → encouraging (adj): mang tính khích lệ → discourage (v): làm nản lòng."
 ],
 [
  "RELY",
  "Rely (v): dựa vào, tin cậy (+ on) → reliable (adj): đáng tin cậy → unreliable (adj): không đáng tin cậy → reliability (n): độ tin cậy → reliably (adv): một cách đáng tin cậy."
 ],
 [
  "ELECTRIC",
  "Electric (adj): thuộc về điện, chạy bằng điện → electrical (adj): liên quan đến điện (hệ thống, an toàn điện) → electricity (n): điện năng → electrician (n): thợ điện."
 ],
 [
  "BALANCE",
  "Balance (n, v): sự cân bằng; làm cho cân bằng → balanced (adj): cân bằng, hợp lý → unbalanced (adj): mất cân bằng."
 ],
 [
  "APPROPRIATE",
  "Appropriate (adj): phù hợp, thích hợp → inappropriate (adj): không phù hợp → appropriately (adv): một cách phù hợp, thích đáng."
 ],
 [
  "REDUCE",
  "Reduce (v): cắt giảm, làm giảm → reduction (n): sự cắt giảm."
 ],
 [
  "ANXIOUS",
  "Anxious (adj): lo lắng, bồn chồn → anxiety (n): sự lo âu, nỗi lo lắng → anxiously (adv): một cách lo âu."
 ],
 [
  "RESPONSIBLE",
  "Responsible (adj): có trách nhiệm → irresponsible (adj): vô trách nhiệm → responsibility (n): trách nhiệm → responsibly (adv): một cách có trách nhiệm → irresponsibly (adv): một cách vô trách nhiệm."
 ],
 [
  "PARTICIPATE",
  "Participate (v): tham gia → participation (n): sự tham gia → participant (n): người tham gia."
 ],
 [
  "CONGEST",
  "Congest (v): làm tắc nghẽn → congested (adj): bị tắc nghẽn, đông nghịt → congestion (n): sự tắc nghẽn (giao thông)."
 ],
 [
  "PRIORITY",
  "Priority (n): sự ưu tiên, việc ưu tiên → prioritize / prioritise (v): ưu tiên, sắp xếp theo thứ tự ưu tiên → prioritizing / prioritising (V-ing): việc ưu tiên."
 ],
 [
  "POLLUTE",
  "Pollute (v): gây ô nhiễm → pollution (n): sự ô nhiễm → pollutant (n): chất gây ô nhiễm → polluted (adj): bị ô nhiễm."
 ],
 [
  "IMPROVE",
  "Improve (v): cải thiện, nâng cấp → improvement (n): sự cải thiện."
 ],
 [
  "CRIME",
  "Crime (n): tội phạm, hành vi phạm tội → criminal (n, adj): kẻ phạm tội (n); thuộc về tội phạm, phạm pháp (adj)."
 ],
 [
  "PRESERVE",
  "Preserve (v): bảo tồn, giữ gìn → preservation (n): sự bảo tồn."
 ],
 [
  "DISTRACT",
  "Distract (v): làm xao nhãng → distraction (n): sự xao nhãng, yếu tố gây xao nhãng → distracted (adj): bị phân tâm → distracting (adj): gây phân tâm."
 ],
 [
  "EQUIP",
  "Equip (v): trang bị → equipment (n): trang thiết bị (không đếm được) → equipped (adj): được trang bị."
 ],
 [
  "PRODUCE",
  "Produce (v): sản xuất, tạo ra → product (n): sản phẩm → production (n): sự sản xuất → productive (adj): năng suất, hiệu quả → productively (adv): một cách năng suất, có hiệu quả."
 ],
 [
  "AVOID",
  "Avoid (v): tránh, né tránh → avoidance (n): sự né tránh → avoidable (adj): có thể tránh được → unavoidable (adj): không thể tránh khỏi → avoiding (V-ing): việc tránh."
 ]
];

// Câu mẫu theo cấu trúc đề tuyển sinh 10 TP.HCM: 17 cột theo sheet CauHoi (cột P = Unit, Q = Chủ đề)
var SAMPLE_QUESTIONS_NEW = [
 [
  41,
  "Choose the word whose underlined part is pronounced differently from the others.",
  "",
  "C",
  "",
  "'bread' đọc /e/ (/bred/), còn beach /iː/, team /iː/, meat /iː/.",
  "Chọn từ có phần gạch chân phát âm khác với các từ còn lại.",
  "",
  "",
  "Trắc nghiệm",
  "b**ea**ch",
  "t**ea**m",
  "br**ea**d",
  "m**ea**t",
  "",
  "Ôn tổng hợp",
  "Phát âm"
 ],
 [
  42,
  "Choose the word whose underlined part is pronounced differently from the others.",
  "",
  "C",
  "",
  "Đuôi -ed đọc /ɪd/ sau âm /t/ và /d/ (needed /ˈniːdɪd/); còn lại đọc /t/ (helped, watched, stopped).",
  "Chọn từ có phần gạch chân phát âm khác.",
  "",
  "",
  "Trắc nghiệm",
  "help**ed**",
  "watch**ed**",
  "need**ed**",
  "stopp**ed**",
  "",
  "Ôn tổng hợp",
  "Phát âm"
 ],
 [
  43,
  "Choose the word that has a different stress pattern from the others.",
  "",
  "C",
  "",
  "'difficult' nhấn âm 1 (/ˈdɪfɪkəlt/); important, tradition, attention nhấn âm 2.",
  "Chọn từ có trọng âm khác với các từ còn lại.",
  "",
  "",
  "Trắc nghiệm",
  "important",
  "tradition",
  "difficult",
  "attention",
  "",
  "Ôn tổng hợp",
  "Trọng âm"
 ],
 [
  44,
  "Choose the word that has a different stress pattern from the others.",
  "",
  "D",
  "",
  "'decision' nhấn âm 2 (/dɪˈsɪʒn/); beautiful, tourism, mountain nhấn âm 1.",
  "Chọn từ có trọng âm khác với các từ còn lại.",
  "",
  "",
  "Trắc nghiệm",
  "beautiful",
  "tourism",
  "mountain",
  "decision",
  "",
  "Ôn tổng hợp",
  "Trọng âm"
 ],
 [
  45,
  "If I ________ more free time, I would join the volunteer club.",
  "",
  "B",
  "",
  "Câu điều kiện loại 2 (giả định không có thật ở hiện tại): If + S + V-ed, S + would + V.",
  "Nếu tôi có nhiều thời gian rảnh hơn, tôi sẽ tham gia câu lạc bộ tình nguyện.",
  "",
  "",
  "Trắc nghiệm",
  "have",
  "had",
  "will have",
  "would have",
  "",
  "Ôn tổng hợp",
  "Câu điều kiện"
 ],
 [
  46,
  "The Temple of Literature ________ in 1070.",
  "",
  "B",
  "",
  "Câu bị động ở quá khứ đơn (có mốc thời gian 1070): was/were + V3.",
  "Văn Miếu được xây dựng vào năm 1070.",
  "",
  "",
  "Trắc nghiệm",
  "built",
  "was built",
  "is built",
  "has built",
  "",
  "Ôn tổng hợp",
  "Câu bị động"
 ],
 [
  47,
  "My brother said that he ________ to the zoo the next day.",
  "",
  "C",
  "",
  "Câu tường thuật: lùi thì. 'will go' -> 'would go'; 'tomorrow' -> 'the next day'.",
  "Anh trai tôi nói rằng ngày hôm sau anh ấy sẽ đi sở thú.",
  "",
  "",
  "Trắc nghiệm",
  "goes",
  "went",
  "would go",
  "has gone",
  "",
  "Ôn tổng hợp",
  "Câu tường thuật"
 ],
 [
  48,
  "The city is ________ than it was ten years ago.",
  "",
  "B",
  "",
  "So sánh hơn với tính từ dài: more + adj + than.",
  "Thành phố đông đúc hơn so với mười năm trước.",
  "",
  "",
  "Trắc nghiệm",
  "crowded",
  "more crowded",
  "most crowded",
  "the more crowded",
  "",
  "Ôn tổng hợp",
  "So sánh"
 ],
 [
  49,
  "She is the girl ________ won the first prize in the speaking contest.",
  "",
  "C",
  "",
  "Đại từ quan hệ thay cho người và làm chủ ngữ của mệnh đề: who.",
  "Cô ấy là cô gái đã giành giải nhất cuộc thi nói.",
  "",
  "",
  "Trắc nghiệm",
  "whose",
  "which",
  "who",
  "whom",
  "",
  "Ôn tổng hợp",
  "Mệnh đề quan hệ"
 ],
 [
  50,
  "I wish I ________ speak English as fluently as my teacher.",
  "",
  "B",
  "",
  "Câu ước ở hiện tại: wish + S + V-ed/could + V (điều không có thật).",
  "Tôi ước mình có thể nói tiếng Anh trôi chảy như cô giáo.",
  "",
  "",
  "Trắc nghiệm",
  "can",
  "could",
  "will",
  "would",
  "",
  "Ôn tổng hợp",
  "Câu ước"
 ],
 [
  51,
  "We have lived in this neighbourhood ________ 2015.",
  "",
  "B",
  "",
  "Hiện tại hoàn thành: since + mốc thời gian (2015); for + khoảng thời gian.",
  "Chúng tôi sống ở khu phố này từ năm 2015.",
  "",
  "",
  "Trắc nghiệm",
  "for",
  "since",
  "during",
  "in",
  "",
  "Ôn tổng hợp",
  "Thì hiện tại hoàn thành"
 ],
 [
  52,
  "He is ________ tired that he can't keep his eyes open.",
  "",
  "C",
  "",
  "Cấu trúc so + adj + that + mệnh đề (quá ... đến nỗi ...).",
  "Anh ấy mệt đến nỗi không thể mở mắt nổi.",
  "",
  "",
  "Trắc nghiệm",
  "such",
  "too",
  "so",
  "enough",
  "",
  "Ôn tổng hợp",
  "Cấu trúc so...that / such...that"
 ],
 [
  53,
  "Students are encouraged ________ part in community activities.",
  "",
  "B",
  "",
  "Cấu trúc bị động: be encouraged + to-V.",
  "Học sinh được khuyến khích tham gia các hoạt động cộng đồng.",
  "",
  "",
  "Trắc nghiệm",
  "take",
  "to take",
  "taking",
  "took",
  "",
  "Ôn tổng hợp",
  "Câu bị động"
 ],
 [
  54,
  "________ the heavy rain, the football match went on as planned.",
  "",
  "C",
  "",
  "Despite + cụm danh từ (the heavy rain). Although cần một mệnh đề đầy đủ.",
  "Bất chấp cơn mưa lớn, trận bóng đá vẫn diễn ra như kế hoạch.",
  "",
  "",
  "Trắc nghiệm",
  "Although",
  "Because",
  "Despite",
  "However",
  "",
  "Ôn tổng hợp",
  "Liên từ chỉ sự tương phản"
 ],
 [
  55,
  "Please ________ the light when you leave the room.",
  "",
  "A",
  "",
  "turn off = tắt (đèn); turn on = bật; turn up = vặn to; turn into = biến thành.",
  "Hãy tắt đèn khi bạn rời khỏi phòng.",
  "",
  "",
  "Trắc nghiệm",
  "turn off",
  "turn up",
  "turn on",
  "turn into",
  "",
  "Ôn tổng hợp",
  "Cụm động từ"
 ],
 [
  56,
  "Ha Long Bay attracts millions of ________ every year.",
  "",
  "A",
  "",
  "Sau 'millions of' cần danh từ đếm được số nhiều chỉ người: tourists.",
  "Vịnh Hạ Long thu hút hàng triệu khách du lịch mỗi năm.",
  "",
  "",
  "Trắc nghiệm",
  "tourists",
  "tours",
  "tourism",
  "tourist",
  "",
  "Unit 5: Wonders of Viet Nam",
  "Từ vựng"
 ],
 [
  57,
  "— Would you like another slice of cake?  — ________",
  "",
  "A",
  "",
  "Đáp lại lời mời: 'Yes, please.' (nhận lời) hoặc 'No, thanks.' (từ chối).",
  "— Bạn dùng thêm một miếng bánh nữa nhé? — Vâng, xin cảm ơn. Bánh rất ngon.",
  "",
  "",
  "Trắc nghiệm",
  "Yes, please. It's delicious.",
  "No, I wouldn't.",
  "It's my pleasure.",
  "You're welcome.",
  "",
  "Ôn tổng hợp",
  "Giao tiếp"
 ],
 [
  58,
  "— Congratulations on winning the competition!  — ________",
  "",
  "B",
  "",
  "Đáp lại lời chúc mừng: Thank you.",
  "— Chúc mừng bạn đã thắng cuộc thi! — Cảm ơn bạn.",
  "",
  "",
  "Trắc nghiệm",
  "Never mind.",
  "Thank you.",
  "That's a pity.",
  "You're welcome.",
  "",
  "Ôn tổng hợp",
  "Giao tiếp"
 ],
 [
  59,
  "We should ________ more attention to our health.",
  "",
  "A",
  "",
  "Kết hợp từ cố định: pay attention to (chú ý đến).",
  "Chúng ta nên chú ý nhiều hơn đến sức khỏe của mình.",
  "",
  "",
  "Trắc nghiệm",
  "pay",
  "take",
  "make",
  "give",
  "",
  "Ôn tổng hợp",
  "Cụm từ cố định"
 ],
 [
  60,
  "Teenagers often feel ________ when they have too much homework.",
  "",
  "C",
  "",
  "Sau 'feel' cần tính từ chỉ cảm xúc của người: stressed (căng thẳng).",
  "Thanh thiếu niên thường cảm thấy căng thẳng khi có quá nhiều bài tập.",
  "",
  "",
  "Trắc nghiệm",
  "stress",
  "stressful",
  "stressed",
  "stressing",
  "",
  "Unit 3: Teen Stress and Pressure",
  "Từ vựng"
 ],
 [
  61,
  "People grow a lot of rice in the Mekong Delta.",
  "Begin with: A lot of rice ...",
  "A lot of rice is grown in the Mekong Delta.",
  "",
  "Chuyển sang bị động hiện tại đơn: S + is/are + V3 + (by ...). Chủ ngữ 'People' chung chung nên bỏ.",
  "Người ta trồng rất nhiều lúa ở đồng bằng sông Cửu Long.",
  "",
  "",
  "Viết lại câu",
  "",
  "",
  "",
  "",
  "",
  "Ôn tổng hợp",
  "Câu bị động"
 ],
 [
  62,
  "\"I am preparing for my English test,\" Lan said.",
  "Begin with: Lan said that ...",
  "Lan said that she was preparing for her English test.",
  "Lan said she was preparing for her English test.",
  "Câu tường thuật: đổi ngôi (I -> she, my -> her) và lùi thì (am preparing -> was preparing).",
  "Lan nói rằng cô ấy đang chuẩn bị cho bài kiểm tra tiếng Anh.",
  "",
  "",
  "Viết lại câu",
  "",
  "",
  "",
  "",
  "",
  "Ôn tổng hợp",
  "Câu tường thuật"
 ],
 [
  63,
  "Study harder or you will fail the exam.",
  "Begin with: If you ...",
  "If you don't study harder, you will fail the exam.",
  "If you do not study harder, you will fail the exam.",
  "Câu điều kiện loại 1: If + S + V (hiện tại), S + will + V. 'or' = nếu không.",
  "Nếu không học chăm hơn, bạn sẽ trượt kỳ thi.",
  "",
  "",
  "Viết lại câu",
  "",
  "",
  "",
  "",
  "",
  "Ôn tổng hợp",
  "Câu điều kiện"
 ],
 [
  64,
  "The soup was so hot that I couldn't eat it.",
  "Begin with: The soup was too ...",
  "The soup was too hot for me to eat.",
  "",
  "so + adj + that + mệnh đề phủ định = too + adj + (for O) + to-V. Không lặp lại tân ngữ 'it'.",
  "Món súp nóng đến mức tôi không thể ăn.",
  "",
  "",
  "Viết lại câu",
  "",
  "",
  "",
  "",
  "",
  "Unit 7: Recipes and Eating Habits",
  "Cấu trúc too / enough"
 ],
 [
  65,
  "Despite the bad weather, we enjoyed the trip.",
  "Begin with: Although ...",
  "Although the weather was bad, we enjoyed the trip.",
  "",
  "Despite + cụm danh từ = Although + mệnh đề (S + V). 'the bad weather' -> 'the weather was bad'.",
  "Mặc dù thời tiết xấu, chúng tôi vẫn tận hưởng chuyến đi.",
  "",
  "",
  "Viết lại câu",
  "",
  "",
  "",
  "",
  "",
  "Ôn tổng hợp",
  "Liên từ chỉ sự tương phản"
 ],
 [
  66,
  "This book is not as interesting as that one.",
  "Begin with: That book ...",
  "That book is more interesting than this one.",
  "That book is more interesting than this book.",
  "A not as + adj + as B = B + be + more + adj + than A.",
  "Cuốn sách kia thú vị hơn cuốn này.",
  "",
  "",
  "Viết lại câu",
  "",
  "",
  "",
  "",
  "",
  "Ôn tổng hợp",
  "So sánh"
 ],
 [
  67,
  "My grandfather doesn't live in the countryside any more.",
  "Begin with: My grandfather used to ...",
  "My grandfather used to live in the countryside.",
  "",
  "used to + V nguyên mẫu: thói quen/trạng thái trong quá khứ nay không còn.",
  "Ông tôi từng sống ở nông thôn.",
  "",
  "",
  "Viết lại câu",
  "",
  "",
  "",
  "",
  "",
  "Unit 4: Life in the Past",
  "Used to"
 ],
 [
  68,
  "They began learning English three years ago.",
  "Begin with: They have ...",
  "They have learned English for three years.",
  "They have learnt English for three years.|They have been learning English for three years.",
  "Hiện tại hoàn thành: S + have/has + V3 + for + khoảng thời gian.",
  "Họ đã học tiếng Anh được ba năm.",
  "",
  "",
  "Viết lại câu",
  "",
  "",
  "",
  "",
  "",
  "Ôn tổng hợp",
  "Thì hiện tại hoàn thành"
 ],
 [
  69,
  "I'm sorry I don't have a bicycle.",
  "Begin with: I wish ...",
  "I wish I had a bicycle.",
  "",
  "Câu ước ở hiện tại: wish + S + V-ed (trái với sự thật hiện tại).",
  "Tôi ước gì mình có một chiếc xe đạp.",
  "",
  "",
  "Viết lại câu",
  "",
  "",
  "",
  "",
  "",
  "Ôn tổng hợp",
  "Câu ước"
 ],
 [
  70,
  "The room is too small for all of us.",
  "Begin with: The room isn't ...",
  "The room isn't big enough for all of us.",
  "The room is not big enough for all of us.",
  "too + adj = not + adj (trái nghĩa) + enough. 'too small' -> 'not big enough'.",
  "Căn phòng không đủ rộng cho tất cả chúng tôi.",
  "",
  "",
  "Viết lại câu",
  "",
  "",
  "",
  "",
  "",
  "Ôn tổng hợp",
  "Cấu trúc too / enough"
 ],
 [
  71,
  "My parents never ________ me. They only want me to try my best.",
  "",
  "put pressure on",
  "",
  "Dùng cụm 'put pressure on' (gây áp lực lên) trong ví dụ thứ hai của mục từ.",
  "Bố mẹ tôi không bao giờ gây áp lực lên tôi. Họ chỉ muốn tôi cố gắng hết sức.",
  "",
  "",
  "Word entry",
  "",
  "",
  "",
  "",
  "pressure",
  "Unit 3: Teen Stress and Pressure",
  "Cụm từ cố định"
 ],
 [
  72,
  "Some teenagers start smoking because of ________ from their classmates.",
  "",
  "peer pressure",
  "",
  "Cụm 'peer pressure' (áp lực từ bạn bè đồng trang lứa) trong ví dụ thứ ba.",
  "Một số thanh thiếu niên bắt đầu hút thuốc vì áp lực từ bạn cùng lớp.",
  "",
  "",
  "Word entry",
  "",
  "",
  "",
  "",
  "pressure",
  "Unit 3: Teen Stress and Pressure",
  "Cụm từ cố định"
 ],
 [
  73,
  "Drinking enough water every day has a lot of ________.",
  "",
  "health benefits",
  "",
  "Cụm 'health benefits' (lợi ích cho sức khỏe) ở ví dụ đầu tiên.",
  "Uống đủ nước mỗi ngày có rất nhiều lợi ích cho sức khỏe.",
  "",
  "",
  "Word entry",
  "",
  "",
  "",
  "",
  "benefit",
  "Unit 7: Recipes and Eating Habits",
  "Cụm từ cố định"
 ],
 [
  74,
  "Elderly people can ________ taking part in outdoor activities.",
  "",
  "benefit from",
  "",
  "'benefit from + V-ing': được lợi từ việc gì (ví dụ thứ hai).",
  "Người lớn tuổi có thể được lợi từ việc tham gia các hoạt động ngoài trời.",
  "",
  "",
  "Word entry",
  "",
  "",
  "",
  "",
  "benefit",
  "Ôn tổng hợp",
  "Cụm từ cố định"
 ],
 [
  75,
  "Eating too much sugar is a ________ that you should stop.",
  "",
  "bad habit",
  "",
  "Cụm 'bad habit' (thói quen xấu) ở ví dụ đầu; ngữ cảnh 'should stop' cho thấy đây là thói quen xấu.",
  "Ăn quá nhiều đường là một thói quen xấu mà bạn nên bỏ.",
  "",
  "",
  "Word entry",
  "",
  "",
  "",
  "",
  "habit",
  "Unit 7: Recipes and Eating Habits",
  "Cụm từ cố định"
 ],
 [
  76,
  "My teacher advises us to ________ of doing homework right after school.",
  "",
  "develop a habit",
  "develop the habit",
  "Cụm 'develop a habit of + V-ing' (hình thành thói quen) ở ví dụ thứ ba.",
  "Cô giáo khuyên chúng tôi hình thành thói quen làm bài tập ngay sau giờ học.",
  "",
  "",
  "Word entry",
  "",
  "",
  "",
  "",
  "habit",
  "Unit 3: Teen Stress and Pressure",
  "Cụm từ cố định"
 ]
];

// Mục từ điển mẫu cho Word entry: [Từ, IPA, Từ loại, Nghĩa, Ví dụ]
var SAMPLE_ENTRIES = [
 [
  "pressure",
  "/ˈpreʃ.ər/",
  "noun",
  "1. the worried feeling you have when other people expect a lot from you\n2. the force of one thing pressing on another",
  "Many students feel **under pressure** when the exams are coming.\nParents should not **put pressure on** their children to get high marks.\n**Peer pressure** can make teenagers do things they do not want to do.\nTalking to friends can help you **cope with pressure**."
 ],
 [
  "benefit",
  "/ˈben.ɪ.fɪt/",
  "noun, verb",
  "1. (noun) a good or helpful result of something\n2. (verb) to be helped by something",
  "Regular exercise has many **health benefits**.\nTeenagers can **benefit from** joining volunteer clubs.\nThe park was built **for the benefit of** local people.\nLearning a language is good for the mind."
 ],
 [
  "habit",
  "/ˈhæb.ɪt/",
  "noun",
  "1. something that you do regularly or often, without thinking about it",
  "Smoking is a **bad habit** that harms your health.\nFast food has changed the **eating habits** of many young people.\nTeachers should help students **develop a habit** of reading every day.\nChildren learn **good habits** from their parents."
 ]
];


// Hai trang web được nhúng sẵn để chỉ cần mở link /exec là dùng.
var PAGE_STUDENT_ = "<!DOCTYPE html>\n<html lang=\"vi\">\n<head>\n    <meta charset=\"UTF-8\">\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n    <title>Word Form 9 by NMC</title>\n    <style>\n        /* ==========================================================================\n           100% OFFLINE CSS SYSTEM - DÀNH RIÊNG CHO WORD FORM 9 BY NMC\n           ========================================================================== */\n        :root {\n            --bg-dark: #0f172a;\n            --bg-card: #1e293b;\n            --bg-card-hover: #26354d;\n            --bg-input: #0b1120;\n            --text-main: #f8fafc;\n            --text-muted: #94a3b8;\n            --primary: #38bdf8;\n            --primary-dark: #0284c7;\n            --accent: #fbbf24;\n            --accent-dark: #d97706;\n            --success: #22c55e;\n            --success-bg: rgba(34, 197, 94, 0.14);\n            --danger: #f43f5e;\n            --danger-bg: rgba(244, 63, 94, 0.14);\n            --border: rgba(255, 255, 255, 0.12);\n            --radius-lg: 20px;\n            --radius-md: 12px;\n            --radius-sm: 8px;\n            --shadow: 0 10px 30px -5px rgba(0, 0, 0, 0.45);\n            --font-sans: system-ui, -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;\n        }\n\n        * {\n            box-sizing: border-box;\n            margin: 0;\n            padding: 0;\n        }\n\n        body {\n            font-family: var(--font-sans);\n            background: radial-gradient(circle at top, #172554 0%, #0f172a 55%, #090d16 100%);\n            background-attachment: fixed;\n            color: var(--text-main);\n            min-height: 100vh;\n            line-height: 1.6;\n            padding-bottom: 60px;\n        }\n\n        .container {\n            max-width: 1080px;\n            margin: 0 auto;\n            padding: 16px;\n        }\n\n        /* Header */\n        .app-header {\n            background: rgba(30, 41, 59, 0.88);\n            backdrop-filter: blur(10px);\n            border: 1px solid var(--border);\n            border-radius: var(--radius-lg);\n            padding: 16px 22px;\n            margin-bottom: 18px;\n            box-shadow: var(--shadow);\n            display: flex;\n            flex-wrap: wrap;\n            justify-content: space-between;\n            align-items: center;\n            gap: 14px;\n        }\n\n        .brand-box {\n            display: flex;\n            align-items: center;\n            gap: 12px;\n        }\n\n        .brand-logo {\n            width: 46px;\n            height: 46px;\n            border-radius: 14px;\n            background: linear-gradient(135deg, #38bdf8, #6366f1);\n            display: flex;\n            align-items: center;\n            justify-content: center;\n            font-size: 22px;\n            font-weight: 800;\n            color: white;\n            box-shadow: 0 4px 12px rgba(56, 189, 248, 0.35);\n            flex-shrink: 0;\n        }\n\n        .brand-title {\n            font-size: 1.45rem;\n            font-weight: 800;\n            color: #fff;\n            letter-spacing: -0.02em;\n        }\n\n        /* Navigation Tabs */\n        .nav-tabs {\n            display: grid;\n            grid-template-columns: repeat(3, 1fr);\n            gap: 10px;\n            margin-bottom: 18px;\n        }\n\n        .tab-btn {\n            background: rgba(30, 41, 59, 0.75);\n            border: 1px solid var(--border);\n            color: var(--text-muted);\n            padding: 13px 16px;\n            border-radius: var(--radius-md);\n            font-size: 0.95rem;\n            font-weight: 700;\n            cursor: pointer;\n            transition: all 0.2s ease;\n            display: flex;\n            align-items: center;\n            justify-content: center;\n            gap: 8px;\n            text-align: center;\n        }\n\n        .tab-btn:hover {\n            background: var(--bg-card-hover);\n            color: #fff;\n        }\n\n        .tab-btn.active {\n            background: linear-gradient(135deg, #fbbf24, #f59e0b);\n            color: #0f172a;\n            border-color: transparent;\n            box-shadow: 0 4px 15px rgba(245, 158, 11, 0.35);\n        }\n\n        /* Stats Dashboard Bar */\n        .stats-bar {\n            display: grid;\n            grid-template-columns: repeat(4, 1fr);\n            gap: 10px;\n            margin-bottom: 16px;\n        }\n\n        .stat-card {\n            background: rgba(30, 41, 59, 0.78);\n            border: 1px solid var(--border);\n            border-radius: var(--radius-md);\n            padding: 12px 14px;\n            display: flex;\n            flex-direction: column;\n            align-items: center;\n            text-align: center;\n        }\n\n        .stat-label {\n            font-size: 0.78rem;\n            color: var(--text-muted);\n            font-weight: 600;\n            text-transform: uppercase;\n            letter-spacing: 0.04em;\n        }\n\n        .stat-val {\n            font-size: 1.45rem;\n            font-weight: 800;\n            margin-top: 2px;\n        }\n\n        /* Progress Bar */\n        .progress-wrap {\n            width: 100%;\n            background: rgba(255,255,255,0.08);\n            height: 8px;\n            border-radius: 99px;\n            overflow: hidden;\n            margin-bottom: 20px;\n        }\n\n        .progress-fill {\n            height: 100%;\n            background: linear-gradient(90deg, #38bdf8, #22c55e);\n            width: 0%;\n            transition: width 0.35s ease;\n        }\n\n        /* Buttons */\n        .btn {\n            border: none;\n            border-radius: var(--radius-sm);\n            padding: 9px 14px;\n            font-size: 0.88rem;\n            font-weight: 700;\n            cursor: pointer;\n            transition: all 0.15s ease;\n            display: inline-flex;\n            align-items: center;\n            justify-content: center;\n            gap: 6px;\n            font-family: inherit;\n        }\n\n        .btn:active {\n            transform: translateY(1px);\n        }\n\n        .btn-primary {\n            background: var(--primary);\n            color: #0f172a;\n        }\n        .btn-primary:hover {\n            background: #7dd3fc;\n        }\n\n        .btn-accent {\n            background: var(--accent);\n            color: #0f172a;\n        }\n        .btn-accent:hover {\n            background: #fcd34d;\n        }\n\n        .btn-success {\n            background: var(--success);\n            color: #052e16;\n        }\n        .btn-success:hover {\n            background: #4ade80;\n        }\n\n        .btn-ghost {\n            background: rgba(255, 255, 255, 0.09);\n            color: var(--text-main);\n            border: 1px solid var(--border);\n        }\n        .btn-ghost:hover {\n            background: rgba(255, 255, 255, 0.16);\n        }\n\n        .btn-danger-ghost {\n            background: rgba(244, 63, 94, 0.15);\n            color: #fda4af;\n            border: 1px solid rgba(244, 63, 94, 0.3);\n        }\n        .btn-danger-ghost:hover {\n            background: rgba(244, 63, 94, 0.28);\n        }\n\n        /* Main Card Panels */\n        .panel {\n            background: rgba(30, 41, 59, 0.88);\n            border: 1px solid var(--border);\n            border-radius: var(--radius-lg);\n            padding: 24px;\n            box-shadow: var(--shadow);\n        }\n\n        .hidden {\n            display: none !important;\n        }\n\n        /* v10: chọn giọng đọc (US/UK, Nam/Nữ, tốc độ) */\n        .voice-panel { margin: 0 0 14px; padding: 12px 14px; border-radius: 14px; background: rgba(255,255,255,0.05); border: 1px solid var(--border); }\n        .voice-row { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-bottom: 9px; }\n        .voice-row > .vlabel { min-width: 74px; font-size: 0.82rem; font-weight: 700; color: var(--text-muted); }\n        .vopt { padding: 7px 12px; border-radius: 10px; border: 1px solid var(--border); background: transparent; color: var(--text-main); font: inherit; font-size: 0.86rem; font-weight: 600; cursor: pointer; min-height: 38px; }\n        .vopt.on { background: var(--accent); color: #1f1300; border-color: var(--accent); }\n        .voice-info { font-size: 0.8rem; color: var(--text-muted); line-height: 1.45; margin: 2px 0 0; }\n        .voice-info.warn { color: #fde68a; }\n\n        /* v11: cách phát âm (Sayshape) */\n        .ph-modal-card { background: #1e293b; border: 2px solid var(--border); border-radius: var(--radius-lg); max-width: 560px; width: 100%; max-height: 88vh; overflow-y: auto; padding: 20px; box-shadow: var(--shadow); text-align: left; }\n        .ph-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 10px; margin-bottom: 6px; }\n        .ph-head h3 { font-size: 1.12rem; font-weight: 800; }\n        .ph-sub { color: var(--text-muted); font-size: 0.86rem; margin: 0 0 12px; word-break: break-word; }\n        .ph-note { font-size: 0.84rem; color: #fde68a; background: rgba(251,191,36,0.1); border: 1px solid rgba(251,191,36,0.4); border-radius: 10px; padding: 8px 10px; margin-bottom: 10px; line-height: 1.45; }\n        .ph-row { display: flex; gap: 12px; padding: 10px 0; border-top: 1px solid var(--border); }\n        .ph-sym { min-width: 54px; font-size: 1.5rem; font-weight: 800; color: var(--accent); text-align: center; line-height: 1.2; }\n        .ph-title { font-weight: 800; font-size: 0.95rem; margin-bottom: 2px; }\n        .ph-tip { font-size: 0.84rem; color: var(--text-muted); line-height: 1.5; margin-bottom: 6px; }\n        .ph-links { display: flex; flex-wrap: wrap; gap: 6px; }\n        .ph-a { display: inline-block; padding: 6px 11px; border-radius: 10px; border: 1px solid var(--accent); color: var(--accent); font-weight: 700; font-size: 0.82rem; text-decoration: none; min-height: 34px; line-height: 22px; }\n        .ph-a.off { border-color: var(--border); color: var(--text-muted); }\n        .ph-foot { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px; padding-top: 12px; border-top: 1px solid var(--border); }\n\n        /* v8: giám sát rời trang */\n        .ac-bar { position: sticky; top: 6px; z-index: 40; margin: 0 0 12px; padding: 10px 14px; border-radius: 12px; background: rgba(251, 191, 36, 0.14); border: 1px solid rgba(251, 191, 36, 0.5); color: #fde68a; font-size: 0.88rem; font-weight: 600; line-height: 1.45; backdrop-filter: blur(4px); }\n        .ac-bar.bad { background: var(--danger-bg); border-color: var(--danger); color: #fecdd3; }\n        .ac-lock .ws-item, .ac-lock .ws-sentence, .ac-lock .rw-hint { -webkit-user-select: none; user-select: none; -webkit-touch-callout: none; }\n        .ac-lock input, .ac-lock textarea { -webkit-user-select: text; user-select: text; }\n\n        /* ==========================================================================\n           MODE 1: LUYỆN TẬP TỪNG CÂU (SINGLE QUESTION PRACTICE)\n           ========================================================================== */\n        .filter-bar-top {\n            display: flex;\n            justify-content: space-between;\n            align-items: center;\n            margin-bottom: 10px;\n            flex-wrap: wrap;\n            gap: 8px;\n        }\n\n        .question-nav-pills {\n            display: flex;\n            flex-wrap: wrap;\n            gap: 6px;\n            margin-bottom: 18px;\n            max-height: 115px;\n            overflow-y: auto;\n            padding: 8px;\n            background: rgba(15, 23, 42, 0.65);\n            border-radius: var(--radius-md);\n            border: 1px solid var(--border);\n        }\n\n        .ws-item { content-visibility: auto; contain-intrinsic-size: auto 230px; }\n        .q-pill {\n            width: 36px;\n            height: 34px;\n            border-radius: 8px;\n            border: 1px solid var(--border);\n            background: rgba(255,255,255,0.05);\n            color: var(--text-muted);\n            font-weight: 700;\n            font-size: 0.82rem;\n            cursor: pointer;\n            display: flex;\n            align-items: center;\n            justify-content: center;\n            transition: all 0.15s;\n        }\n\n        .q-pill:hover {\n            border-color: var(--primary);\n            color: #fff;\n        }\n\n        .q-pill.current {\n            outline: 2px solid var(--accent);\n            background: rgba(251, 191, 36, 0.22);\n            color: var(--accent);\n        }\n\n        .q-pill.done-correct {\n            background: rgba(34, 197, 94, 0.25);\n            border-color: #22c55e;\n            color: #86efac;\n        }\n\n        .q-pill.done-wrong {\n            background: rgba(244, 63, 94, 0.25);\n            border-color: #f43f5e;\n            color: #fda4af;\n        }\n\n        .focus-card {\n            background: linear-gradient(180deg, rgba(15, 23, 42, 0.92) 0%, rgba(15, 23, 42, 0.75) 100%);\n            border: 2px solid rgba(56, 189, 248, 0.25);\n            border-radius: var(--radius-lg);\n            padding: 26px 22px;\n            margin-bottom: 18px;\n        }\n\n        .badge-row {\n            display: flex;\n            flex-wrap: wrap;\n            justify-content: space-between;\n            align-items: center;\n            gap: 10px;\n            margin-bottom: 16px;\n        }\n\n        .root-badge {\n            background: linear-gradient(135deg, #fbbf24, #f59e0b);\n            color: #0f172a;\n            font-weight: 800;\n            padding: 6px 16px;\n            border-radius: 99px;\n            font-size: 0.95rem;\n            letter-spacing: 0.03em;\n            box-shadow: 0 3px 10px rgba(251, 191, 36, 0.3);\n        }\n\n        .sentence-display {\n            font-size: 1.38rem;\n            font-weight: 700;\n            line-height: 1.65;\n            color: #fff;\n            margin: 18px 0 24px 0;\n            text-align: center;\n        }\n\n        .blank-highlight {\n            display: inline-block;\n            min-width: 120px;\n            border-bottom: 3px dashed var(--accent);\n            color: var(--accent);\n            padding: 0 8px;\n            font-weight: 800;\n        }\n\n        .blank-filled-correct {\n            background: rgba(34, 197, 94, 0.25);\n            border-bottom: 3px solid #22c55e;\n            color: #86efac;\n            padding: 2px 10px;\n            border-radius: 6px;\n        }\n\n        .blank-filled-wrong {\n            background: rgba(244, 63, 94, 0.25);\n            border-bottom: 3px solid #f43f5e;\n            color: #fda4af;\n            padding: 2px 10px;\n            border-radius: 6px;\n        }\n\n        .root-inline {\n            color: var(--accent);\n            font-weight: 800;\n            margin-left: 6px;\n        }\n\n        /* Input Row */\n        .input-action-row {\n            display: flex;\n            flex-wrap: wrap;\n            gap: 10px;\n            max-width: 640px;\n            margin: 0 auto;\n        }\n\n        .answer-input {\n            flex: 1;\n            min-width: 220px;\n            background: #ffffff;\n            color: #0f172a;\n            border: 3px solid transparent;\n            border-radius: var(--radius-md);\n            padding: 13px 18px;\n            font-size: 1.15rem;\n            font-weight: 700;\n            outline: none;\n            transition: all 0.2s;\n        }\n\n        .answer-input:focus {\n            border-color: var(--accent);\n            box-shadow: 0 0 0 4px rgba(251, 191, 36, 0.25);\n        }\n\n        .answer-input.input-correct {\n            background: #dcfce7;\n            border-color: #22c55e;\n            color: #14532d;\n        }\n\n        .answer-input.input-wrong {\n            background: #ffe4e6;\n            border-color: #f43f5e;\n            color: #881337;\n        }\n\n        /* Hints Row */\n        .hints-toolbar {\n            display: flex;\n            flex-wrap: wrap;\n            justify-content: center;\n            gap: 10px;\n            margin-top: 16px;\n        }\n\n        .hint-drawer {\n            margin-top: 14px;\n            background: rgba(56, 189, 248, 0.12);\n            border: 1px solid rgba(56, 189, 248, 0.35);\n            border-radius: var(--radius-md);\n            padding: 12px 16px;\n            font-size: 0.94rem;\n            color: #e0f2fe;\n        }\n\n        /* Explanation & Translation Result Box */\n        .explanation-box {\n            margin-top: 20px;\n            border-radius: var(--radius-md);\n            padding: 18px 20px;\n            border: 2px solid;\n            animation: fadeIn 0.25s ease;\n        }\n\n        @keyframes fadeIn {\n            from { opacity: 0; transform: translateY(6px); }\n            to { opacity: 1; transform: translateY(0); }\n        }\n\n        .explanation-box.correct {\n            background: var(--success-bg);\n            border-color: rgba(34, 197, 94, 0.5);\n        }\n\n        .explanation-box.wrong {\n            background: var(--danger-bg);\n            border-color: rgba(244, 63, 94, 0.5);\n        }\n\n        .exp-header {\n            display: flex;\n            flex-wrap: wrap;\n            justify-content: space-between;\n            align-items: center;\n            gap: 8px;\n            margin-bottom: 12px;\n            padding-bottom: 10px;\n            border-bottom: 1px solid rgba(255,255,255,0.12);\n        }\n\n        .exp-status {\n            font-size: 1.12rem;\n            font-weight: 800;\n        }\n\n        .exp-detail-grid {\n            display: grid;\n            gap: 10px;\n            font-size: 0.96rem;\n        }\n\n        .exp-item {\n            background: rgba(15, 23, 42, 0.6);\n            padding: 11px 14px;\n            border-radius: 10px;\n            border-left: 4px solid var(--primary);\n        }\n\n        .exp-item.translation {\n            border-left-color: var(--accent);\n        }\n\n        .exp-item.family {\n            border-left-color: #a855f7;\n        }\n\n        /* ==========================================================================\n           MODE 2: LÀM ĐỀ TỔNG HỢP (WORKSHEET / FULL TEST MODE)\n           ========================================================================== */\n        .worksheet-controls {\n            display: flex;\n            flex-wrap: wrap;\n            justify-content: space-between;\n            align-items: center;\n            gap: 12px;\n            margin-bottom: 18px;\n            padding-bottom: 14px;\n            border-bottom: 1px solid var(--border);\n        }\n\n        .score-banner {\n            background: linear-gradient(135deg, rgba(34, 197, 94, 0.18), rgba(56, 189, 248, 0.18));\n            border: 2px solid rgba(56, 189, 248, 0.45);\n            border-radius: var(--radius-md);\n            padding: 18px 22px;\n            margin-bottom: 20px;\n            display: flex;\n            flex-wrap: wrap;\n            justify-content: space-between;\n            align-items: center;\n            gap: 14px;\n        }\n\n        .ws-list {\n            display: flex;\n            flex-direction: column;\n            gap: 14px;\n        }\n\n        .ws-item {\n            background: rgba(15, 23, 42, 0.68);\n            border: 1px solid var(--border);\n            border-radius: var(--radius-md);\n            padding: 16px;\n            transition: border-color 0.2s;\n        }\n\n        .ws-item.ws-correct {\n            border-color: rgba(34, 197, 94, 0.65);\n            background: rgba(34, 197, 94, 0.07);\n        }\n\n        .ws-item.ws-wrong {\n            border-color: rgba(244, 63, 94, 0.65);\n            background: rgba(244, 63, 94, 0.07);\n        }\n\n        .ws-top-line {\n            display: flex;\n            flex-wrap: wrap;\n            justify-content: space-between;\n            align-items: center;\n            gap: 10px;\n            margin-bottom: 8px;\n        }\n\n        .ws-q-num {\n            font-weight: 800;\n            color: var(--primary);\n            font-size: 0.95rem;\n        }\n\n        .ws-sentence {\n            font-size: 1.08rem;\n            font-weight: 600;\n            line-height: 1.65;\n            color: #f1f5f9;\n        }\n\n        .ws-inline-input {\n            background: #ffffff;\n            color: #0f172a;\n            border: 2px solid #94a3b8;\n            padding: 5px 12px;\n            border-radius: 8px;\n            font-weight: 700;\n            font-size: 1rem;\n            width: 195px;\n            outline: none;\n            margin: 0 4px;\n            text-align: center;\n        }\n\n        .ws-inline-input:focus {\n            border-color: var(--accent);\n            box-shadow: 0 0 0 3px rgba(251, 191, 36, 0.25);\n        }\n\n        .ws-inline-input.ws-inp-ok {\n            background: #dcfce7;\n            border-color: #22c55e;\n            color: #14532d;\n        }\n\n        .ws-inline-input.ws-inp-bad {\n            background: #ffe4e6;\n            border-color: #f43f5e;\n            color: #881337;\n        }\n\n        .ws-feedback {\n            margin-top: 12px;\n            padding: 12px 14px;\n            border-radius: 10px;\n            background: rgba(15, 23, 42, 0.85);\n            font-size: 0.92rem;\n            border-left: 4px solid var(--primary);\n        }\n\n        /* ==========================================================================\n           MODE 3: SỔ TAY GIA ĐÌNH TỪ (WORD FAMILIES HANDBOOK)\n           ========================================================================== */\n        .search-bar-wrap {\n            display: flex;\n            flex-wrap: wrap;\n            gap: 10px;\n            margin-bottom: 18px;\n        }\n\n        .search-input {\n            flex: 1;\n            min-width: 240px;\n            background: rgba(15, 23, 42, 0.9);\n            border: 1px solid var(--border);\n            color: #fff;\n            padding: 12px 16px;\n            border-radius: var(--radius-md);\n            font-size: 0.98rem;\n            outline: none;\n        }\n\n        .search-input:focus {\n            border-color: var(--primary);\n        }\n\n        .wf-grid {\n            display: grid;\n            grid-template-columns: repeat(auto-fill, minmax(310px, 1fr));\n            gap: 14px;\n        }\n\n        .wf-card {\n            background: rgba(15, 23, 42, 0.78);\n            border: 1px solid var(--border);\n            border-radius: var(--radius-md);\n            padding: 16px;\n            display: flex;\n            flex-direction: column;\n            justify-content: space-between;\n            transition: transform 0.15s, border-color 0.15s;\n        }\n\n        .wf-card:hover {\n            border-color: rgba(56, 189, 248, 0.5);\n            transform: translateY(-2px);\n        }\n\n        .wf-card-header {\n            display: flex;\n            justify-content: space-between;\n            align-items: center;\n            padding-bottom: 10px;\n            margin-bottom: 10px;\n            border-bottom: 1px solid var(--border);\n        }\n\n        .wf-title {\n            font-size: 1.1rem;\n            font-weight: 800;\n            color: var(--accent);\n        }\n\n        .wf-branch {\n            padding: 5px 0 5px 10px;\n            border-left: 2px solid rgba(56, 189, 248, 0.4);\n            margin-bottom: 6px;\n            font-size: 0.92rem;\n        }\n\n        /* Custom Confirmation Modal */\n        .modal-backdrop {\n            position: fixed;\n            inset: 0;\n            background: rgba(0, 0, 0, 0.75);\n            display: flex;\n            align-items: center;\n            justify-content: center;\n            z-index: 999;\n            padding: 16px;\n        }\n\n        .modal-card {\n            background: #1e293b;\n            border: 2px solid var(--border);\n            border-radius: var(--radius-lg);\n            max-width: 420px;\n            width: 100%;\n            padding: 24px;\n            text-align: center;\n            box-shadow: var(--shadow);\n        }\n\n        @media (max-width: 768px) {\n            .nav-tabs {\n                grid-template-columns: 1fr;\n            }\n            .stats-bar {\n                grid-template-columns: repeat(2, 1fr);\n            }\n            .sentence-display {\n                font-size: 1.16rem;\n            }\n            .panel {\n                padding: 16px;\n            }\n        }\n\n        /* ===== BỔ SUNG: đăng nhập + chip học sinh (phần còn lại của giao diện giữ nguyên file gốc) ===== */\n        .login-screen { min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 16px; }\n        .login-card { width: 100%; max-width: 440px; background: rgba(30, 41, 59, 0.92); border: 1px solid var(--border); border-radius: var(--radius-lg); padding: 28px 24px; box-shadow: var(--shadow); }\n        .login-field { display: flex; flex-direction: column; gap: 6px; margin-bottom: 14px; }\n        .login-field label { font-size: 0.8rem; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; }\n        .login-input { width: 100%; background: #ffffff; color: #0f172a; border: 3px solid transparent; border-radius: var(--radius-md); padding: 12px 14px; font-size: 1.05rem; font-weight: 700; outline: none; font-family: inherit; }\n        .login-input:focus { border-color: var(--accent); box-shadow: 0 0 0 4px rgba(251, 191, 36, 0.25); }\n        .login-error { background: var(--danger-bg); border: 1px solid rgba(244, 63, 94, 0.5); color: #fecdd3; border-radius: 10px; padding: 10px 12px; font-size: 0.9rem; margin-bottom: 14px; }\n        .login-note { margin-top: 14px; font-size: 0.8rem; color: var(--text-muted); text-align: center; }\n        .student-chip { display: inline-flex; align-items: center; gap: 6px; flex-wrap: wrap; background: rgba(56, 189, 248, 0.14); border: 1px solid rgba(56, 189, 248, 0.4); color: #e0f2fe; border-radius: 99px; padding: 6px 14px; font-size: 0.85rem; font-weight: 700; cursor: pointer; font-family: inherit; }\n        .student-chip .sync { font-weight: 600; font-size: 0.78rem; color: #86efac; }\n        .student-chip .sync.busy { color: #fcd34d; }\n        .student-chip .sync.err { color: #fda4af; }\n        /* ===== MƯỢT TRÊN ĐIỆN THOẠI: cùng giao diện, nhưng nền cố định dùng lớp riêng (không vẽ lại khi cuộn) ===== */\n        html { background: #0f172a; }\n        body { background: transparent; background-attachment: scroll; }\n        body::before { content: \"\"; position: fixed; inset: 0; z-index: -1; pointer-events: none; background: radial-gradient(circle at top, #172554 0%, #0f172a 55%, #090d16 100%); }\n        button, .q-pill, .tab-btn { touch-action: manipulation; -webkit-tap-highlight-color: transparent; }\n        @media (max-width: 900px) {\n            .app-header { backdrop-filter: none; -webkit-backdrop-filter: none; }\n            .search-input { font-size: 16px; }\n        }\n        @media (prefers-reduced-motion: reduce) { .explanation-box { animation: none; } }\n        .header-right { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }\n\n        /* ===== v5: 4 dạng bài ===== */\n        .hl { color: var(--accent); text-decoration: underline; text-underline-offset: 3px; font-weight: 800; }\n        /* ===== v6: lọc theo Unit / Chủ đề / từ khóa ===== */\n        .scope-bar { background: rgba(30,41,59,0.88); border: 1px solid var(--border); border-radius: var(--radius-lg); padding: 14px 16px; margin-bottom: 16px; }\n        .scope-head { display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-bottom:10px; }\n        .scope-title { font-weight:800; color:#fff; font-size:0.98rem; }\n        .scope-search { flex:1 1 220px; min-width:180px; background: rgba(15,23,42,0.9); color: var(--text-main); border: 1px solid var(--border); border-radius: 10px; padding: 8px 12px; font-size: 0.9rem; font-family: inherit; }\n        .scope-search:focus { outline:none; border-color: var(--primary); }\n        .scope-sum { font-size:0.82rem; color: var(--text-muted); width:100%; }\n        .scope-sum b { color: var(--accent); }\n        .unit-cards { display:grid; grid-template-columns: repeat(auto-fill, minmax(158px, 1fr)); gap:8px; max-height: 230px; overflow-y:auto; padding:2px; }\n        .unit-card { text-align:left; background: rgba(15,23,42,0.75); border:2px solid var(--border); color: var(--text-main); border-radius:12px; padding:9px 11px; cursor:pointer; font-family:inherit; transition:border-color .15s, background .15s; display:flex; flex-direction:column; gap:3px; }\n        .unit-card:hover { border-color: var(--primary); background: rgba(56,189,248,0.08); }\n        .unit-card.active { border-color: var(--accent); background: rgba(251,191,36,0.12); }\n        .unit-card .u-no { font-weight:800; font-size:0.92rem; color:#fff; }\n        .unit-card .u-name { font-size:0.76rem; color: var(--text-muted); line-height:1.25; min-height:1.9em; }\n        .unit-card .u-meta { font-size:0.74rem; color: var(--text-muted); display:flex; justify-content:space-between; gap:6px; }\n        .unit-card .u-bar { height:5px; border-radius:99px; background: rgba(255,255,255,0.1); overflow:hidden; }\n        .unit-card .u-bar i { display:block; height:100%; background: var(--primary); }\n        .unit-card .u-bar i.ok { background:#4ade80; }\n        .topic-chips { display:flex; flex-wrap:wrap; gap:6px; margin-top:10px; max-height:96px; overflow-y:auto; }\n        .t-chip { background: rgba(15,23,42,0.75); border:1px solid var(--border); color: var(--text-muted); border-radius:99px; padding:4px 11px; font-size:0.8rem; font-weight:700; cursor:pointer; font-family:inherit; }\n        .t-chip:hover { border-color: var(--primary); color: var(--text-main); }\n        .t-chip.active { border-color: var(--accent); background: rgba(251,191,36,0.14); color: var(--accent); }\n        .t-chip small { opacity:.75; font-weight:600; margin-left:3px; }\n        .scope-empty { text-align:center; padding:28px 12px; color: var(--text-muted); font-weight:700; border:1px dashed var(--border); border-radius: var(--radius-md); margin-bottom:16px; }\n        .q-unit-tag { display:inline-block; font-size:0.72rem; font-weight:700; padding:2px 9px; border-radius:99px; margin-left:6px; background:rgba(255,255,255,0.06); color: var(--text-muted); border:1px solid var(--border); vertical-align:middle; }\n        .type-select { background: rgba(15,23,42,0.9); color: var(--text-main); border: 1px solid var(--border); border-radius: 8px; padding: 6px 10px; font-size: 0.8rem; font-weight: 700; }\n        .type-tag { display:inline-block; font-size:0.72rem; font-weight:800; padding:2px 9px; border-radius:99px; margin-left:6px; background:rgba(56,189,248,0.16); color:var(--primary); border:1px solid rgba(56,189,248,0.35); vertical-align:middle; }\n        .type-tag.t-trac_nghiem { background:rgba(167,139,250,0.16); color:#c4b5fd; border-color:rgba(167,139,250,0.4); }\n        .type-tag.t-viet_lai { background:rgba(244,114,182,0.16); color:#f9a8d4; border-color:rgba(244,114,182,0.4); }\n        .type-tag.t-word_entry { background:rgba(52,211,153,0.16); color:#6ee7b7; border-color:rgba(52,211,153,0.4); }\n        .opt-list { display:flex; flex-direction:column; gap:9px; max-width:640px; margin:0 auto 14px auto; }\n        .opt-btn { display:flex; align-items:center; gap:12px; width:100%; text-align:left; background:rgba(15,23,42,0.75); color:var(--text-main); border:2px solid var(--border); border-radius:12px; padding:11px 14px; font-size:1.05rem; font-weight:600; cursor:pointer; font-family:inherit; transition:border-color .15s, background .15s; }\n        .opt-btn:hover:not(:disabled) { border-color: var(--primary); background: rgba(56,189,248,0.08); }\n        .opt-btn.selected { border-color: var(--accent); background: rgba(251,191,36,0.12); }\n        .opt-btn.opt-correct { border-color:#22c55e; background:rgba(34,197,94,0.18); }\n        .opt-btn.opt-wrong { border-color:#f43f5e; background:rgba(244,63,94,0.18); }\n        .opt-letter { flex:0 0 auto; width:30px; height:30px; border-radius:50%; background:rgba(255,255,255,0.1); display:flex; align-items:center; justify-content:center; font-weight:800; color:var(--primary); }\n        .opt-btn.selected .opt-letter { background:var(--accent); color:#0f172a; }\n        .opt-btn.opt-correct .opt-letter { background:#22c55e; color:#052e16; }\n        .opt-btn.opt-wrong .opt-letter { background:#f43f5e; color:#fff; }\n        .ws-opts { display:flex; flex-direction:column; gap:7px; margin-top:10px; }\n        .ws-opts .opt-btn { font-size:0.98rem; padding:8px 12px; }\n        .ws-opts input[type=radio] { position:absolute; opacity:0; pointer-events:none; }\n        .rw-hint { margin-top:8px; font-size:0.98rem; font-weight:700; color:var(--accent); }\n        .answer-textarea { flex:1; min-width:220px; min-height:76px; background:#fff; color:#0f172a; border:3px solid transparent; border-radius:12px; padding:12px 16px; font-size:1.1rem; font-weight:700; outline:none; font-family:inherit; resize:vertical; }\n        .answer-textarea:focus { border-color: var(--accent); box-shadow: 0 0 0 4px rgba(251,191,36,0.25); }\n        .answer-textarea.input-correct { background:#dcfce7; border-color:#22c55e; color:#14532d; }\n        .answer-textarea.input-wrong { background:#ffe4e6; border-color:#f43f5e; color:#881337; }\n        .ws-textarea { width:100%; margin-top:10px; min-height:64px; background:#fff; color:#0f172a; border:2px solid #94a3b8; border-radius:8px; padding:8px 12px; font-size:1rem; font-weight:700; font-family:inherit; resize:vertical; outline:none; }\n        .ws-textarea:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(251,191,36,0.25); }\n        .ws-textarea.ws-inp-ok { background:#dcfce7; border-color:#22c55e; color:#14532d; }\n        .ws-textarea.ws-inp-bad { background:#ffe4e6; border-color:#f43f5e; color:#881337; }\n        .entry-card { background:rgba(15,23,42,0.8); border:1px solid rgba(52,211,153,0.45); border-left:5px solid #34d399; border-radius:12px; padding:12px 16px; margin:0 auto 16px auto; max-width:760px; text-align:left; font-size:0.95rem; line-height:1.6; }\n        .entry-head { display:flex; flex-wrap:wrap; align-items:baseline; gap:10px; margin-bottom:4px; }\n        .entry-word { font-size:1.3rem; font-weight:800; color:#fff; }\n        .entry-ipa { color:var(--text-muted); font-weight:600; }\n        .entry-pos { color:#6ee7b7; font-style:italic; font-weight:700; }\n        .entry-card ol, .entry-card ul { margin:4px 0 6px 22px; }\n        .entry-card li { margin:2px 0; color:#e2e8f0; }\n        .entry-card details > summary { cursor:pointer; font-weight:800; color:#6ee7b7; }\n        .de-title { margin:22px 0 10px 0; font-size:1.05rem; font-weight:800; color:#6ee7b7; }\n        @media (max-width: 768px) { .sentence-display { font-size: 1.18rem; } .opt-btn { font-size: 1rem; } }\n    </style>\n</head>\n<body>\n\n    <!-- ===== MÀN HÌNH ĐĂNG NHẬP: chỉ học sinh có trong danh sách lớp mới vào được ===== -->\n    <div id=\"login-screen\" class=\"login-screen\">\n        <form class=\"login-card\" onsubmit=\"return submitLogin(event)\" autocomplete=\"off\">\n            <div style=\"text-align:center; margin-bottom:20px;\">\n                <div class=\"brand-logo\" style=\"margin:0 auto 10px; width:56px; height:56px; font-size:28px;\">🎓</div>\n                <h1 class=\"brand-title\">Word Form 9 by NMC</h1>\n                <p style=\"color:var(--text-muted); font-size:0.92rem; margin-top:4px;\">Đăng nhập để làm bài. Kết quả tự động gửi cho giáo viên.</p>\n            </div>\n            <div id=\"login-error\" class=\"login-error hidden\"></div>\n            <div class=\"login-field\">\n                <label for=\"login-class\">Lớp</label>\n                <select id=\"login-class\" class=\"login-input\"><option value=\"\">Đang tải danh sách lớp...</option></select>\n            </div>\n            <div class=\"login-field\">\n                <label for=\"login-name\">Họ và tên</label>\n                <input id=\"login-name\" class=\"login-input\" type=\"text\" placeholder=\"Gõ đầy đủ họ và tên (có dấu)\" autocomplete=\"off\" spellcheck=\"false\">\n            </div>\n            <div id=\"login-code-wrap\" class=\"login-field hidden\">\n                <label for=\"login-code\">Mã học sinh</label>\n                <input id=\"login-code\" class=\"login-input\" type=\"text\" placeholder=\"Mã do giáo viên cấp\" autocomplete=\"off\" spellcheck=\"false\">\n            </div>\n            <button id=\"login-btn\" type=\"submit\" class=\"btn btn-accent\" style=\"width:100%; padding:14px; font-size:1.05rem;\">Vào làm bài ➜</button>\n            <p class=\"login-note\">Chọn đúng lớp và gõ đúng họ tên (có dấu) như trong danh sách lớp.<br>Chỉ học sinh có tên trong danh sách mới vào được.</p>\n        </form>\n    </div>\n\n    <div id=\"notice-modal\" class=\"modal-backdrop hidden\">\n        <div class=\"modal-card\">\n            <p id=\"notice-text\" style=\"font-size:1rem; margin-bottom:20px;\"></p>\n            <div style=\"display:flex; gap:10px; justify-content:center;\">\n                <button id=\"notice-cancel\" class=\"btn btn-ghost hidden\" style=\"padding:10px 18px;\">Hủy</button>\n                <button id=\"notice-ok\" class=\"btn btn-accent\" style=\"padding:10px 22px;\">OK</button>\n            </div>\n        </div>\n    </div>\n\n    <div id=\"ph-modal\" class=\"modal-backdrop hidden\" onclick=\"if(event.target===this)closePhonetics()\">\n        <div class=\"ph-modal-card\" role=\"dialog\" aria-modal=\"true\" aria-labelledby=\"ph-title\">\n            <div class=\"ph-head\"><h3 id=\"ph-title\">👄 Cách phát âm</h3>\n                <button type=\"button\" class=\"btn btn-ghost\" style=\"padding:4px 10px;\" onclick=\"closePhonetics()\" aria-label=\"Đóng\">✕</button></div>\n            <p id=\"ph-sub\" class=\"ph-sub\"></p>\n            <div id=\"ph-body\"></div>\n            <div id=\"ph-foot\" class=\"ph-foot\"></div>\n        </div>\n    </div>\n\n    <div id=\"loading-overlay\" class=\"modal-backdrop hidden\">\n        <div class=\"modal-card\">\n            <div style=\"font-size:2.2rem; margin-bottom:8px;\">⏳</div>\n            <p id=\"loading-text\" style=\"font-size:1rem;\">Đang tải câu hỏi...</p>\n        </div>\n    </div>\n\n    <div class=\"container hidden\" id=\"app-root\">\n        <!-- HEADER (KHÔNG CÓ SUBTITLE VÀ KHÔNG CÓ NÚT TẢI OFFLINE) -->\n        <header class=\"app-header\">\n            <div class=\"brand-box\">\n                <div class=\"brand-logo\">🎓</div>\n                <h1 class=\"brand-title\">Word Form 9 by NMC</h1>\n            </div>\n            <div class=\"header-right\">\n                <button type=\"button\" class=\"student-chip\" onclick=\"logoutStudent()\" title=\"Bấm để đăng xuất / đổi tài khoản\">👤 <span id=\"student-chip-text\"></span> <span id=\"sync-state\" class=\"sync\">✔</span></button>\n                <button type=\"button\" id=\"voice-toggle\" class=\"btn btn-ghost\" onclick=\"toggleVoicePanel()\" title=\"Chọn giọng đọc: Mỹ/Anh, Nam/Nữ, tốc độ\" aria-expanded=\"false\" aria-controls=\"voice-panel\">🔊 Giọng đọc</button>\n                <button onclick=\"openResetModal()\" class=\"btn btn-danger-ghost\" title=\"Xóa lịch sử làm bài để làm lại từ đầu\">\n                    🔄 Làm lại từ đầu\n                </button>\n            </div>\n        </header>\n\n        <div id=\"app-error\" class=\"login-error hidden\"></div>\n        <div id=\"voice-panel\" class=\"voice-panel hidden\" role=\"group\" aria-label=\"Chọn giọng đọc\">\n            <div class=\"voice-row\"><span class=\"vlabel\">Giọng vùng</span>\n                <button type=\"button\" class=\"vopt\" data-vk=\"acc\" data-vv=\"us\" onclick=\"setVoicePref('acc','us')\">🇺🇸 Mỹ (US)</button>\n                <button type=\"button\" class=\"vopt\" data-vk=\"acc\" data-vv=\"uk\" onclick=\"setVoicePref('acc','uk')\">🇬🇧 Anh (UK)</button></div>\n            <div class=\"voice-row\"><span class=\"vlabel\">Giọng đọc</span>\n                <button type=\"button\" class=\"vopt\" data-vk=\"sex\" data-vv=\"f\" onclick=\"setVoicePref('sex','f')\">👩 Nữ</button>\n                <button type=\"button\" class=\"vopt\" data-vk=\"sex\" data-vv=\"m\" onclick=\"setVoicePref('sex','m')\">👨 Nam</button></div>\n            <div class=\"voice-row\"><span class=\"vlabel\">Tốc độ</span>\n                <button type=\"button\" class=\"vopt\" data-vk=\"spd\" data-vv=\"0.55\" onclick=\"setVoicePref('spd','0.55')\">🐢 Rất chậm</button>\n                <button type=\"button\" class=\"vopt\" data-vk=\"spd\" data-vv=\"0.75\" onclick=\"setVoicePref('spd','0.75')\">🚶 Chậm</button>\n                <button type=\"button\" class=\"vopt\" data-vk=\"spd\" data-vv=\"0.92\" onclick=\"setVoicePref('spd','0.92')\">🏃 Bình thường</button></div>\n            <div class=\"voice-row\"><span class=\"vlabel\">Lặp lại</span>\n                <button type=\"button\" class=\"vopt\" data-vk=\"rep\" data-vv=\"0\" onclick=\"setVoicePref('rep','0')\">Đọc 1 lần</button>\n                <button type=\"button\" class=\"vopt\" data-vk=\"rep\" data-vv=\"2\" onclick=\"setVoicePref('rep','2')\">🔁 Đọc 2 lần (lần 2 chậm hơn)</button></div>\n            <div class=\"voice-row\"><button type=\"button\" class=\"btn btn-ghost\" style=\"padding:6px 12px;\" onclick=\"testVoice()\">▶ Nghe thử</button>\n                <a href=\"#\" class=\"vopt\" style=\"text-decoration:none; display:inline-flex; align-items:center;\" onclick=\"return openPhatamPage(event)\">📚 Âm khó với người Việt</a></div>\n            <p id=\"voice-info\" class=\"voice-info\"></p>\n        </div>\n        <div id=\"ac-bar\" class=\"ac-bar hidden\" role=\"status\" aria-live=\"polite\"></div>\n\n        <!-- STATS BAR -->\n        <div class=\"stats-bar\">\n            <div class=\"stat-card\">\n                <span class=\"stat-label\">Tiến độ hoàn thành</span>\n                <span class=\"stat-val\" style=\"color: var(--primary);\" id=\"stat-progress\">0 / 0</span>\n            </div>\n            <div class=\"stat-card\">\n                <span class=\"stat-label\">Số câu đúng</span>\n                <span class=\"stat-val\" style=\"color: var(--success);\" id=\"stat-correct\">0</span>\n            </div>\n            <div class=\"stat-card\">\n                <span class=\"stat-label\">Câu làm sai</span>\n                <span class=\"stat-val\" style=\"color: var(--danger);\" id=\"stat-wrong\">0</span>\n            </div>\n            <div class=\"stat-card\">\n                <span class=\"stat-label\">🔥 Chuỗi đúng liên tiếp</span>\n                <span class=\"stat-val\" style=\"color: var(--accent);\" id=\"stat-streak\">0</span>\n            </div>\n        </div>\n\n        <!-- PROGRESS BAR -->\n        <div class=\"progress-wrap\" id=\"progress-wrap\" title=\"Tiến độ hoàn thành\">\n            <div id=\"main-progress-bar\" class=\"progress-fill\"></div>\n        </div>\n\n        <!-- NAVIGATION TABS -->\n        <nav class=\"nav-tabs\">\n            <button onclick=\"switchMode('flashcard')\" id=\"tab-flashcard\" class=\"tab-btn active\">\n                ✍️ Chế độ 1: Luyện tập từng câu\n            </button>\n            <button onclick=\"switchMode('worksheet')\" id=\"tab-worksheet\" class=\"tab-btn\">\n                📋 Chế độ 2: Làm đề tổng hợp\n            </button>\n            <button onclick=\"switchMode('handbook')\" id=\"tab-handbook\" class=\"tab-btn\">\n                📖 Chế độ 3: Sổ tay (Gia đình từ &amp; Từ điển)\n            </button>\n        </nav>\n\n        <!-- THANH CHỌN PHẦN ÔN TẬP: lọc theo Unit / Chủ đề / từ khóa (áp dụng cho cả 3 chế độ) -->\n        <div id=\"scope-bar\" class=\"scope-bar hidden\">\n            <div class=\"scope-head\">\n                <span class=\"scope-title\">🗂️ Chọn phần ôn tập</span>\n                <input type=\"search\" id=\"scope-search\" class=\"scope-search\" placeholder=\"🔍 Tìm câu theo từ khóa, Unit, chủ đề (VD: stress, Unit 3, câu bị động)...\" oninput=\"onScopeSearch()\" autocomplete=\"off\">\n                <button id=\"scope-clear\" onclick=\"clearScope()\" class=\"btn btn-ghost hidden\" style=\"padding:6px 12px; font-size:0.8rem;\">✖ Bỏ lọc</button>\n                <span id=\"scope-summary\" class=\"scope-sum\"></span>\n            </div>\n            <div id=\"unit-cards\" class=\"unit-cards\"></div>\n            <div id=\"topic-chips\" class=\"topic-chips\"></div>\n        </div>\n\n        <!-- ==========================================================================\n             CHẾ ĐỘ 1: LUYỆN TẬP TỪNG CÂU (KHÔNG HIỆN GỢI Ý TỪ LOẠI TRƯỚC KHI TRẢ LỜI)\n             ========================================================================== -->\n        <section id=\"section-flashcard\" class=\"panel\">\n            <div class=\"filter-bar-top\">\n                <div style=\"display:flex; gap:6px; flex-wrap:wrap;\">\n                    <button onclick=\"setFlashcardFilter('all')\" id=\"fc-filter-all\" class=\"btn btn-accent\" style=\"padding:6px 12px; font-size:0.8rem;\">\n                        Tất cả (0)\n                    </button>\n                    <button onclick=\"setFlashcardFilter('unanswered')\" id=\"fc-filter-unanswered\" class=\"btn btn-ghost\" style=\"padding:6px 12px; font-size:0.8rem;\">\n                        Chưa làm\n                    </button>\n                    <button onclick=\"setFlashcardFilter('wrong')\" id=\"fc-filter-wrong\" class=\"btn btn-ghost\" style=\"padding:6px 12px; font-size:0.8rem;\">\n                        Câu làm sai\n                    </button>\n                </div>\n                <div style=\"display:flex; gap:6px; align-items:center; flex-wrap:wrap;\">\n                    <select id=\"fc-type-filter\" class=\"type-select hidden\" onchange=\"onFcTypeChange()\" aria-label=\"Lọc theo dạng bài\"></select>\n                    <button onclick=\"shuffleFlashcards()\" class=\"btn btn-ghost\" style=\"padding:6px 12px; font-size:0.8rem;\">\n                        🔀 Trộn câu ngẫu nhiên\n                    </button>\n                </div>\n            </div>\n\n            <!-- Danh sách 40 ô số câu hỏi -->\n            <div id=\"fc-pills-container\" class=\"question-nav-pills\"></div>\n\n            <div id=\"fc-empty\" class=\"scope-empty hidden\">Không có câu nào khớp bộ lọc. Hãy chọn Unit / chủ đề khác hoặc bấm \"Bỏ lọc\".</div>\n            <!-- Khung câu hỏi chính -->\n            <div class=\"focus-card\" id=\"fc-card\">\n                <div class=\"badge-row\">\n                    <span id=\"fc-q-number\" style=\"font-weight:800; color:var(--primary); font-size:1.05rem;\">CÂU 1 / 0</span>\n                    <div style=\"display:flex; align-items:center; gap:8px;\">\n                        <button onclick=\"speakCurrentSentence(false)\" class=\"btn btn-ghost\" style=\"padding:6px 12px;\" title=\"Nghe phát âm câu tiếng Anh\">\n                            🔊 Đọc câu\n                        </button>\n                        <button onclick=\"speakCurrentSentence(false, true)\" class=\"btn btn-ghost\" style=\"padding:6px 12px;\" title=\"Nghe chậm\">🐢 Chậm</button>\n                        <button onclick=\"openPhonetics('card')\" class=\"btn btn-ghost\" style=\"padding:6px 12px;\" title=\"Xem miệng và lưỡi đặt thế nào (Sayshape)\">👄 Cách phát âm</button>\n                        <span class=\"root-badge\" id=\"fc-badge\"><span id=\"fc-root-word\">(ATTRACT)</span></span>\n                    </div>\n                </div>\n\n                <!-- Câu tiếng Anh + Từ gốc in hoa trong ngoặc ở cuối câu -->\n                <div id=\"fc-entry\"></div>\n                <div id=\"fc-sentence\" class=\"sentence-display\"></div>\n                <div id=\"fc-options\" class=\"opt-list hidden\"></div>\n\n                <!-- Ô nhập đáp án Tự luận -->\n                <div class=\"input-action-row\">\n                    <input type=\"text\" id=\"fc-input\" class=\"answer-input\" placeholder=\"Gõ dạng từ đúng vào đây...\" autocomplete=\"off\" spellcheck=\"false\" onkeydown=\"if(event.key==='Enter') handleFlashcardEnter()\">\n                    <textarea id=\"fc-textarea\" class=\"answer-textarea hidden\" placeholder=\"Gõ câu viết lại vào đây...\" autocomplete=\"off\" spellcheck=\"false\" onkeydown=\"if(event.key==='Enter' && !event.shiftKey){ event.preventDefault(); handleFlashcardEnter(); }\"></textarea>\n                    <button id=\"fc-submit-btn\" onclick=\"checkFlashcardAnswer()\" class=\"btn btn-accent\" style=\"padding: 13px 22px; font-size: 1.02rem;\">\n                        Kiểm tra ✔️\n                    </button>\n                </div>\n\n                <!-- Chỉ có nút Xem Gia đình từ và Xem đáp án (TUYỆT ĐỐI KHÔNG CÓ GỢI Ý TỪ LOẠI) -->\n                <div class=\"hints-toolbar\">\n                    <button id=\"fc-fam-btn\" onclick=\"toggleFamilyHint()\" class=\"btn btn-ghost\" style=\"font-size:0.84rem;\">\n                        💡 Xem Gia đình từ (Word Family)\n                    </button>\n                    <button onclick=\"revealFlashcardAnswer()\" class=\"btn btn-ghost\" style=\"font-size:0.84rem; color:#fcd34d;\">\n                        👀 Xem đáp án\n                    </button>\n                </div>\n\n                <!-- Ngăn hiển thị Gia đình từ -->\n                <div id=\"fc-hint-family\" class=\"hint-drawer hidden\"></div>\n\n                <!-- Hộp Giải thích & Dịch nghĩa sau khi trả lời -->\n                <div id=\"fc-explanation-box\" class=\"explanation-box hidden\">\n                    <div class=\"exp-header\">\n                        <span id=\"fc-exp-status\" class=\"exp-status\">🎉 Chính xác!</span>\n                        <button onclick=\"speakCurrentSentence(true)\" class=\"btn btn-ghost\" style=\"padding:5px 12px; font-size:0.82rem;\">\n                            🔊 Nghe phát âm câu\n                        </button>\n                    </div>\n                    <div class=\"exp-detail-grid\">\n                        <div class=\"exp-item\">\n                            <strong id=\"fc-exp-full-label\">✅ Câu hoàn chỉnh:</strong>\n                            <span id=\"fc-exp-full-sentence\" style=\"color:#fef08a; font-weight:700;\"></span>\n                        </div>\n                        <div class=\"exp-item\">\n                            <strong>📌 Giải thích ngữ pháp:</strong>\n                            <span id=\"fc-exp-grammar\"></span>\n                        </div>\n                        <div class=\"exp-item translation\">\n                            <strong>🇻🇳 Dịch nghĩa:</strong>\n                            <span id=\"fc-exp-translation\"></span>\n                        </div>\n                        <div class=\"exp-item\" id=\"fc-exp-entry-wrap\" style=\"display:none;\">\n                            <strong>📖 Trong mục từ điển:</strong>\n                            <span id=\"fc-exp-entry\"></span>\n                        </div>\n                        <div class=\"exp-item family\" id=\"fc-exp-family-wrap\">\n                            <strong>🌳 Gia đình từ (Word Family):</strong>\n                            <span id=\"fc-exp-family\"></span>\n                        </div>\n                    </div>\n                </div>\n            </div>\n\n            <!-- Thanh điều hướng Câu trước / Câu tiếp -->\n            <div id=\"fc-nav-row\" style=\"display:flex; justify-content:space-between; align-items:center; gap:10px; flex-wrap:wrap;\">\n                <button onclick=\"prevFlashcard()\" class=\"btn btn-ghost\" style=\"padding:11px 18px;\">\n                    ⬅️ Câu trước\n                </button>\n                <span style=\"font-size:0.85rem; color:var(--text-muted);\">\n                    Mẹo: Nhấn phím <strong>Enter</strong> để kiểm tra nhanh và nhấn <strong>Enter</strong> lần nữa để sang câu tiếp theo.\n                </span>\n                <button onclick=\"nextFlashcard()\" class=\"btn btn-primary\" style=\"padding:11px 22px;\">\n                    Câu tiếp theo ➡️\n                </button>\n            </div>\n        </section>\n\n        <!-- ==========================================================================\n             CHẾ ĐỘ 2: LÀM ĐỀ TỔNG HỢP (WORKSHEET MODE: 10, 20 HOẶC 40 CÂU)\n             ========================================================================== -->\n        <section id=\"section-worksheet\" class=\"panel hidden\">\n            <div class=\"worksheet-controls\">\n                <div style=\"display:flex; align-items:center; gap:8px; flex-wrap:wrap;\">\n                    <span style=\"font-size:0.9rem; font-weight:700; color:var(--text-muted);\">Chọn số lượng câu:</span>\n                    <button onclick=\"generateWorksheet(10)\" id=\"ws-size-10\" class=\"btn btn-ghost\">🎲 10 câu ngẫu nhiên</button>\n                    <button onclick=\"generateWorksheet(20)\" id=\"ws-size-20\" class=\"btn btn-ghost\">🎲 20 câu ngẫu nhiên</button>\n                    <button onclick=\"generateWorksheet(1000)\" id=\"ws-size-all\" class=\"btn btn-accent\">📋 Tất cả 0 câu</button>\n                    <select id=\"ws-type-filter\" class=\"type-select hidden\" onchange=\"generateWorksheet(appState.wsQuizSize)\" aria-label=\"Lọc theo dạng bài\"></select>\n                </div>\n                <div>\n                    <button onclick=\"generateWorksheet(appState.wsQuizSize)\" class=\"btn btn-ghost\">\n                        🔄 Tạo đề mới\n                    </button>\n                </div>\n            </div>\n\n            <!-- Khung kết quả điểm sau khi bấm Nộp bài -->\n            <div id=\"ws-score-banner\" class=\"score-banner hidden\">\n                <div>\n                    <div style=\"font-size:0.85rem; color:var(--text-muted); font-weight:700; text-transform:uppercase;\">Kết quả bài làm</div>\n                    <div id=\"ws-score-title\" style=\"font-size:1.6rem; font-weight:800; color:#4ade80;\">Điểm: 10.0 / 10</div>\n                </div>\n                <div id=\"ws-score-detail\" style=\"font-size:1.05rem; font-weight:700; color:#f8fafc;\">\n                    Đúng 0 / 0 câu (0%)\n                </div>\n            </div>\n\n            <!-- Danh sách câu hỏi trên 1 trang -->\n            <div id=\"worksheet-list\" class=\"ws-list\"></div>\n\n            <!-- Nút Nộp bài & Chấm điểm ở cuối trang -->\n            <div style=\"margin-top: 24px; text-align: center; padding-top: 16px; border-top: 1px solid var(--border);\">\n                <button onclick=\"submitWorksheetQuiz()\" class=\"btn btn-success\" style=\"padding: 14px 32px; font-size: 1.08rem; box-shadow: 0 6px 20px rgba(34, 197, 94, 0.35);\">\n                    ✅ Nộp bài &amp; Chấm điểm\n                </button>\n            </div>\n        </section>\n\n        <!-- ==========================================================================\n             CHẾ ĐỘ 3: SỔ TAY GIA ĐÌNH TỪ (WORD FAMILIES HANDBOOK)\n             ========================================================================== -->\n        <section id=\"section-handbook\" class=\"panel hidden\">\n            <div class=\"search-bar-wrap\">\n                <input type=\"text\" id=\"wf-search-input\" class=\"search-input\" placeholder=\"🔍 Tìm kiếm từ tiếng Anh hoặc nghĩa tiếng Việt (VD: attract, pressure, căng thẳng, tắc nghẽn, ưu tiên)...\" oninput=\"renderWordFamilies()\">\n            </div>\n            <div id=\"hb-scope-note\" class=\"scope-sum hidden\" style=\"margin-bottom:10px;\"></div>\n            <div id=\"wf-grid-container\" class=\"wf-grid\"></div>\n            <div id=\"de-title\" class=\"de-title hidden\">📚 Mục từ điển (dùng cho dạng Word entry)</div>\n            <div id=\"de-grid-container\" class=\"wf-grid\"></div>\n        </section>\n    </div>\n\n    <!-- MODAL XÁC NHẬN LÀM LẠI TỪ ĐẦU -->\n    <div id=\"reset-modal\" class=\"modal-backdrop hidden\">\n        <div class=\"modal-card\">\n            <div style=\"font-size:2.5rem; margin-bottom:8px;\">🔄</div>\n            <h3 style=\"font-size:1.25rem; font-weight:800; margin-bottom:8px;\">Làm lại từ đầu?</h3>\n            <p style=\"font-size:0.92rem; color:var(--text-muted); margin-bottom:20px;\">\n                Toàn bộ lịch sử làm bài của <span class=\"n-q\">0</span> câu của Word Form 9 sẽ được xóa để bạn bắt đầu lượt ôn tập mới.\n            </p>\n            <div style=\"display:flex; gap:10px; justify-content:center;\">\n                <button onclick=\"closeResetModal()\" class=\"btn btn-ghost\" style=\"padding:10px 18px;\">Hủy</button>\n                <button onclick=\"confirmResetAll()\" class=\"btn btn-danger-ghost\" style=\"background:#f43f5e; color:#fff; padding:10px 18px;\">Đồng ý xóa &amp; Làm lại</button>\n            </div>\n        </div>\n    </div>\n\n    <script>\n        // ==========================================================================\n        // 1. DỮ LIỆU: GIA ĐÌNH TỪ & CÂU HỎI WORD FORM (tải từ giáo viên)\n        // ==========================================================================\n        // Dữ liệu được tải từ máy chủ sau khi học sinh đăng nhập (giáo viên nhập trên trang giáo viên)\n        let wordFamilies = [];\n        let questionsData = [];\n        let entriesData = [];\n\n        // ==========================================================================\n        // 2. QUẢN LÝ TRẠNG THÁI & LƯU TRỮ LOCALSTORAGE (KEY: wf9_nmc_progress_v1)\n        // ==========================================================================\n        let STORAGE_KEY = \"wf9_nmc_progress_v1\";\n\n        let appState = {\n            currentMode: \"flashcard\",\n            fcFilter: \"all\",\n            currentOrder: questionsData.map(q => q.id),\n            currentIndex: 0,\n            streak: 0,\n            // answers[id] = { userText: \"\", status: \"correct\" | \"wrong\" | \"revealed\" }\n            answers: {},\n            // Worksheet Mode State\n            wsQuizSize: 1000,\n            wsQuestionIds: questionsData.map(q => q.id),\n            wsSubmitted: false,\n            // v6: bộ lọc phần ôn tập (Unit / Chủ đề / từ khóa)\n            scope: { unit: \"\", topic: \"\", q: \"\" }\n        };\n\n        function loadState() {\n            try {\n                const raw = localStorage.getItem(STORAGE_KEY);\n                if (raw) {\n                    const parsed = JSON.parse(raw);\n                    if (parsed && typeof parsed === \"object\") {\n                        appState.answers = parsed.answers || {};\n                        appState.streak = parsed.streak || 0;\n                        appState.currentIndex = parsed.currentIndex || 0;\n                    }\n                }\n            } catch (e) {}\n        }\n\n        function saveState() {\n            saveStateLocal();\n            queueSync();\n        }\n\n        function saveStateLocal() {\n            try {\n                localStorage.setItem(STORAGE_KEY, JSON.stringify({\n                    answers: appState.answers,\n                    streak: appState.streak,\n                    currentIndex: appState.currentIndex\n                }));\n            } catch (e) {}\n        }\n\n        // ==========================================================================\n        // 3. ÂM THANH PHẢN HỒI & PHÁT ÂM TIẾNG ANH OFFLINE\n        // ==========================================================================\n        const AudioCtx = window.AudioContext || window.webkitAudioContext;\n        let audioCtx = null;\n\n        function playFeedbackTone(isCorrect) {\n            try {\n                if (!audioCtx) audioCtx = new AudioCtx();\n                const now = audioCtx.currentTime;\n                const osc = audioCtx.createOscillator();\n                const gain = audioCtx.createGain();\n                osc.connect(gain);\n                gain.connect(audioCtx.destination);\n\n                if (isCorrect) {\n                    osc.type = \"triangle\";\n                    osc.frequency.setValueAtTime(523.25, now);\n                    osc.frequency.setValueAtTime(659.25, now + 0.09);\n                    osc.frequency.setValueAtTime(783.99, now + 0.18);\n                    gain.gain.setValueAtTime(0.2, now);\n                    gain.gain.exponentialRampToValueAtTime(0.001, now + 0.38);\n                    osc.start(now);\n                    osc.stop(now + 0.38);\n                } else {\n                    osc.type = \"sawtooth\";\n                    osc.frequency.setValueAtTime(210, now);\n                    osc.frequency.setValueAtTime(155, now + 0.14);\n                    gain.gain.setValueAtTime(0.18, now);\n                    gain.gain.exponentialRampToValueAtTime(0.001, now + 0.35);\n                    osc.start(now);\n                    osc.stop(now + 0.35);\n                }\n            } catch (e) {}\n        }\n\n        // ----- Giọng đọc: Mỹ/Anh, Nam/Nữ, tốc độ (lưu trên từng thiết bị) -----\n        const VOICE_KEY = \"wf9_nmc_voice_v1\";\n        let voicePref = { acc: \"us\", sex: \"f\", spd: \"0.75\", rep: \"0\" };\n        (function loadVoicePref() {\n            try {\n                const o = JSON.parse(localStorage.getItem(VOICE_KEY) || \"null\");\n                if (o && (o.acc === \"us\" || o.acc === \"uk\")) voicePref.acc = o.acc;\n                if (o && (o.sex === \"f\" || o.sex === \"m\")) voicePref.sex = o.sex;\n                if (o && [\"0.55\", \"0.75\", \"0.92\"].indexOf(String(o.spd)) >= 0) voicePref.spd = String(o.spd);\n                if (o && (String(o.rep) === \"0\" || String(o.rep) === \"2\")) voicePref.rep = String(o.rep);\n            } catch (e) {}\n        })();\n        const VOICE_FEMALE = /female|woman|samantha|karen|moira|tessa|serena|kate|susan|hazel|zira|victoria|allison|ava\\b|fiona|veena|nicky|joanna|salli|emma|amy|jenny|aria|libby|sonia|catherine|jessa|ivy|kendra|kimberly|stephanie|martha|libby|maisie/i;\n        const VOICE_MALE = /\\bmale\\b|\\bman\\b|daniel|alex\\b|fred|aaron|david|mark\\b|george|james|ryan|guy\\b|oliver|arthur|thomas|rishi|lee\\b|gordon|brian|matthew|joey|justin|kevin|russell|eric|roger|christopher|davis|liam|ethan/i;\n        function voiceGender(v) {\n            const n = (v.name || \"\");\n            if (/female|woman/i.test(n)) return \"f\";\n            if (/google us english/i.test(n)) return \"f\";   // giọng Chrome mặc định là giọng nữ\n            if (VOICE_MALE.test(n) && !VOICE_FEMALE.test(n)) return \"m\";\n            if (VOICE_FEMALE.test(n)) return \"f\";\n            if (/-(gbb|gbd|gbc|usd|iom|ioc|iog|iol|iob)-/i.test(v.voiceURI || \"\")) return \"?\";\n            return \"?\";\n        }\n        function voiceLang(v) { return String(v.lang || \"\").replace(\"_\", \"-\").toLowerCase(); }\n        function pickVoice() {\n            let list = [];\n            try { list = window.speechSynthesis.getVoices() || []; } catch (e) {}\n            const want = voicePref.acc === \"uk\" ? \"en-gb\" : \"en-us\";\n            let pool = list.filter(v => voiceLang(v) === want);\n            if (!pool.length) pool = list.filter(v => voiceLang(v).indexOf(want) === 0);\n            let exact = pool.filter(v => voiceGender(v) === voicePref.sex);\n            const unknown = pool.filter(v => voiceGender(v) === \"?\");\n            const rank = v => (/enhanced|premium|natural|neural|siri/i.test(v.name) ? 0 : 1) + (v.localService ? 0 : 0.5);\n            const best = a => a.slice().sort((x, y) => rank(x) - rank(y))[0] || null;\n            let v = best(exact), note = \"\";\n            if (!v && unknown.length) { v = best(unknown); note = \"unknown\"; }\n            if (!v && pool.length) { v = best(pool); note = \"gender\"; }\n            if (!v) { const any = list.filter(x => voiceLang(x).indexOf(\"en\") === 0); if (any.length) { v = best(any); note = \"accent\"; } }\n            return { voice: v, note: note, total: list.length };\n        }\n        function refreshVoiceUI() {\n            try {\n                document.querySelectorAll(\"#voice-panel .vopt\").forEach(b => b.classList.toggle(\"on\", String(voicePref[b.dataset.vk]) === b.dataset.vv));\n                const info = document.getElementById(\"voice-info\");\n                if (!info) return;\n                const r = pickVoice();\n                const accName = voicePref.acc === \"uk\" ? \"Anh (UK)\" : \"Mỹ (US)\";\n                const sexName = voicePref.sex === \"m\" ? \"Nam\" : \"Nữ\";\n                info.classList.remove(\"warn\");\n                if (!('speechSynthesis' in window)) { info.textContent = \"Trình duyệt này không hỗ trợ đọc thành tiếng.\"; info.classList.add(\"warn\"); return; }\n                if (!r.voice) { info.textContent = r.total ? \"Máy chưa có giọng tiếng Anh. Hãy thử trình duyệt khác (Chrome/Safari).\" : \"Đang tải danh sách giọng của máy… nếu vẫn trống, hãy bấm Nghe thử một lần.\"; info.classList.add(\"warn\"); return; }\n                let msg = \"Đang dùng: \" + r.voice.name + \" (\" + r.voice.lang + \")\";\n                if (r.note === \"gender\") { msg += \" — máy này chưa có giọng \" + sexName + \" \" + accName + \", đang dùng giọng gần nhất.\"; info.classList.add(\"warn\"); }\n                else if (r.note === \"accent\") { msg += \" — máy này chưa có giọng \" + accName + \", đang dùng giọng tiếng Anh gần nhất.\"; info.classList.add(\"warn\"); }\n                else if (r.note === \"unknown\") { msg += \" — (máy không ghi rõ Nam/Nữ nên có thể khác lựa chọn).\"; }\n                info.textContent = msg;\n            } catch (e) {}\n        }\n        function setVoicePref(k, v) {\n            voicePref[k] = String(v);\n            try { localStorage.setItem(VOICE_KEY, JSON.stringify(voicePref)); } catch (e) {}\n            refreshVoiceUI();\n            testVoice();\n        }\n        function toggleVoicePanel() {\n            const p = document.getElementById(\"voice-panel\"), b = document.getElementById(\"voice-toggle\");\n            if (!p) return;\n            const show = p.classList.contains(\"hidden\");\n            p.classList.toggle(\"hidden\", !show);\n            if (b) b.setAttribute(\"aria-expanded\", show ? \"true\" : \"false\");\n            if (show) refreshVoiceUI();\n        }\n        function testVoice() { speakText(\"The orchestra will give a performance at the concert hall on Saturday.\"); }\n        try { if ('speechSynthesis' in window) window.speechSynthesis.addEventListener(\"voiceschanged\", refreshVoiceUI); } catch (e) {}\n\n        let speakToken = 0;\n        function speakText(text, opts) {\n            if (!('speechSynthesis' in window)) return;\n            opts = opts || {};\n            const tok = ++speakToken;\n            try {\n                window.speechSynthesis.cancel();\n                const r = pickVoice();\n                const base = opts.rate || parseFloat(voicePref.spd) || 0.75;\n                const say = (rate, done) => {\n                    const utter = new SpeechSynthesisUtterance(text);\n                    if (r.voice) { utter.voice = r.voice; utter.lang = r.voice.lang; }\n                    else utter.lang = voicePref.acc === \"uk\" ? \"en-GB\" : \"en-US\";\n                    utter.rate = rate;\n                    if (done) utter.onend = done;\n                    window.speechSynthesis.speak(utter);\n                };\n                const twice = voicePref.rep === \"2\" && !opts.once;\n                say(base, twice ? () => {\n                    if (tok !== speakToken) return;\n                    setTimeout(() => { if (tok === speakToken) say(Math.max(0.45, base * 0.75), null); }, 700);\n                } : null);\n            } catch (e) {}\n        }\n\n        // ----- Xem cách phát âm trên Sayshape (sayshape.dqnotes.com) -----\n        const SS_BASE = \"https://sayshape.dqnotes.com/\";\n        const SS_KEY = { \"iː\": \"fleece\", \"ɪ\": \"kit\", \"ʊ\": \"foot\", \"uː\": \"goose\", \"e\": \"dress\", \"ə\": \"comma\", \"ɜː\": \"nurse\", \"ɔː\": \"thought\", \"æ\": \"trap\", \"ʌ\": \"strut\", \"ɑː\": \"palm\", \"ɒ\": \"lot\", \"ɪə\": \"near\", \"eɪ\": \"face\", \"ʊə\": \"cure\", \"ɔɪ\": \"choice\", \"əʊ\": \"goat\", \"eə\": \"square\", \"aɪ\": \"price\", \"aʊ\": \"mouth\", \"p\": \"p\", \"b\": \"b\", \"t\": \"t\", \"d\": \"d\", \"tʃ\": \"ch\", \"dʒ\": \"j\", \"k\": \"k\", \"g\": \"g\", \"f\": \"f\", \"v\": \"v\", \"θ\": \"th\", \"ð\": \"dh\", \"s\": \"s\", \"z\": \"z\", \"ʃ\": \"sh\", \"ʒ\": \"zh\", \"m\": \"m\", \"n\": \"n\", \"ŋ\": \"ng\", \"h\": \"h\", \"l\": \"l\", \"r\": \"r\", \"w\": \"w\", \"j\": \"y\" };\n        function ssUrl(sym) {\n            const k = SS_KEY[sym];\n            return k ? SS_BASE + \"?sound=\" + k + \"&accent=\" + (voicePref.acc === \"uk\" ? \"uk\" : \"us\") + \"&lang=vi&l1=vi\" : \"\";\n        }\n        // các âm người Việt hay gặp khó (viết lại bằng lời của lớp, tham khảo gợi ý của Sayshape)\n        const PH_INFO = {\n            th: { sym: \"/θ/\", title: \"th không rung (think, three)\", links: [\"θ\"], tip: \"Đặt đầu lưỡi chạm nhẹ sau răng cửa trên, thổi hơi qua khe nhỏ, KHÔNG rung cổ họng. Đừng đọc thành t, s hay f (think ≠ sink, three ≠ free).\" },\n            dh: { sym: \"/ð/\", title: \"th có rung (this, mother)\", links: [\"ð\"], tip: \"Cùng vị trí lưỡi với /θ/ nhưng CÓ rung cổ họng (đặt tay lên cổ sẽ thấy rung). Đừng đọc thành d hay z (they ≠ day).\" },\n            sh: { sym: \"/ʃ/\", title: \"sh (she, ship, invitation)\", links: [\"ʃ\"], tip: \"Chu môi nhẹ, lưỡi lùi hơn /s/, thổi hơi xì dài. Phân biệt she ≠ see, ship ≠ sip.\" },\n            zh: { sym: \"/ʒ/\", title: \"âm /ʒ/ (measure, usual, vision)\", links: [\"ʒ\"], tip: \"Giống /ʃ/ nhưng có rung cổ họng. Đừng đọc thành z hay gi.\" },\n            fs: { sym: \"/s/ /z/\", title: \"s, z ở cuối từ (cats, plays, musicians)\", links: [\"s\", \"z\"], tip: \"Tiếng Việt không có từ kết thúc bằng âm xì nên s cuối từ rất dễ bị nuốt. Hãy thổi hơi xì rõ ở cuối từ.\" },\n            fl: { sym: \"/l/\", title: \"l ở cuối từ (feel, school, people)\", links: [\"l\"], tip: \"Giữ đầu lưỡi chạm lợi sau răng trên cho đến hết âm. Đừng bỏ âm l hay đọc thành \\\"iu\\\".\" },\n            fv: { sym: \"/v/\", title: \"v ở cuối từ (love, five, have)\", links: [\"v\"], tip: \"Răng trên chạm môi dưới và giữ rung cổ đến hết từ. Đừng bỏ âm v.\" },\n            fd: { sym: \"/d/ /b/ /ɡ/\", title: \"d, b, g ở cuối từ (bed, bag, big)\", links: [\"d\", \"b\", \"g\"], tip: \"Kéo dài nguyên âm một chút và giữ âm có rung. Đừng biến thành t, p, k (bed ≠ bet, bag ≠ back).\" },\n            vow: { sym: \"/iː/ /ɪ/\", title: \"iː dài và ɪ ngắn (sheep, ship)\", links: [\"iː\", \"ɪ\"], tip: \"/iː/ dài, môi dàn, lưỡi cao; /ɪ/ ngắn và thả lỏng hơn. Phân biệt seat ≠ sit, leave ≠ live.\" },\n            pk: { sym: \"/p/ /t/ /k/\", title: \"p, t, k đầu từ (pen, ten, key)\", links: [\"p\", \"t\", \"k\"], tip: \"Ở đầu từ có hơi bật nhẹ (đặt tay trước miệng sẽ thấy hơi thổi ra). Thiếu hơi bật thì pen nghe như Ben, coat nghe như goat.\" }\n        };\n        const PH_ORDER = [\"th\", \"dh\", \"sh\", \"zh\", \"fs\", \"fl\", \"fv\", \"fd\", \"vow\", \"pk\"];\n        const PH_DH = new Set(\"this that these those they them their theirs then there therefore than though although other others another mother father brother brothers sister weather together rather either neither clothes breathe smooth bathe with without within whether gather feather leather further farther northern southern bother\".split(\" \"));\n        const PH_SH = new Set(\"machine chef brochure moustache parachute champagne chandelier\".split(\" \"));\n        function phWordIds(w) {\n            const ids = [];\n            if (w === \"the\" || w.length < 2) return ids;\n            if (/th/.test(w) && !/^(thomas|thames|thai|thailand|asthma)/.test(w)) ids.push(PH_DH.has(w) ? \"dh\" : \"th\");\n            if (/sh/.test(w) || /(ti|ci|ssi)(on|al|ous|an)/.test(w) || PH_SH.has(w)) ids.push(\"sh\");\n            if (/[aeiou]sion/.test(w) || /(ea|ei)sure/.test(w) || /^(usual|usually|casual|visual|garage|beige|genre)/.test(w)) ids.push(\"zh\");\n            if (w.length >= 3 && /(s|ce|se|x|z)$/.test(w)) ids.push(\"fs\");\n            if (/(l|ll|le)$/.test(w) && w.length >= 3) ids.push(\"fl\");\n            if (/(v|ve)$/.test(w) && w.length >= 3) ids.push(\"fv\");\n            if (/[^n]?(d|b|g)$/.test(w) && !/ng$/.test(w) && w.length >= 3) ids.push(\"fd\");\n            if (/^(p|t(?!h)|k|c[aouklr])/.test(w) && w.length >= 3) ids.push(\"pk\");\n            return ids;\n        }\n        function phIpaSymbols(ipa) {\n            const re = /tʃ|dʒ|iː|uː|ɜː|ɔː|ɑː|ɪə|eɪ|ʊə|ɔɪ|əʊ|oʊ|eə|aɪ|aʊ|[a-zA-Zɪʊæʌɒθðʃʒŋɡɚɝɔəɑ]/g;\n            return String(ipa || \"\").replace(/[ˈˌ.\\/\\[\\]()\\s]/g, \"\").match(re) || [];\n        }\n        function phIpaIds(ipa) {\n            const sy = phIpaSymbols(ipa), ids = [];\n            if (!sy.length) return ids;\n            if (sy.indexOf(\"θ\") >= 0) ids.push(\"th\");\n            if (sy.indexOf(\"ð\") >= 0) ids.push(\"dh\");\n            if (sy.indexOf(\"ʃ\") >= 0) ids.push(\"sh\");\n            if (sy.indexOf(\"ʒ\") >= 0) ids.push(\"zh\");\n            const last = sy[sy.length - 1], first = sy[0];\n            if (last === \"s\" || last === \"z\") ids.push(\"fs\");\n            if (last === \"l\") ids.push(\"fl\");\n            if (last === \"v\") ids.push(\"fv\");\n            if (last === \"d\" || last === \"b\" || last === \"g\" || last === \"ɡ\") ids.push(\"fd\");\n            if (sy.indexOf(\"iː\") >= 0 || sy.indexOf(\"ɪ\") >= 0) ids.push(\"vow\");\n            if (first === \"p\" || first === \"t\" || first === \"k\") ids.push(\"pk\");\n            return ids;\n        }\n        // trả về [{id, words:[...]}] theo thứ tự ưu tiên, tối đa 6 âm\n        function phAnalyse(text, ipa) {\n            const found = {};\n            const add = (id, w) => { (found[id] = found[id] || []); if (found[id].indexOf(w) < 0 && found[id].length < 3) found[id].push(w); };\n            const words = (String(text || \"\").toLowerCase().match(/[a-z]+(?:'[a-z]+)?/g) || []).map(w => w.replace(/'.*$/, \"\"));\n            if (ipa && words.length === 1) phIpaIds(ipa).forEach(id => add(id, words[0]));\n            else words.forEach(w => phWordIds(w).forEach(id => add(id, w)));\n            return PH_ORDER.filter(id => found[id]).slice(0, 6).map(id => ({ id: id, words: found[id] }));\n        }\n        function phCardText() {\n            const q = getCurrentQuestion();\n            if (!q) return \"\";\n            const record = appState.answers[q.id];\n            const t = qType(q);\n            if (t === \"viet_lai\" || (t === \"trac_nghiem\" && !hasBlank(q))) return plain(q.sentence);\n            return (record && record.status) ? completedText(q) : plain(q.sentence).replace(BLANK, \" \");\n        }\n        function phatamUrl() { return (typeof API_URL === \"string\" && /^https:/.test(API_URL)) ? API_URL + \"?page=phatam\" : \"\"; }\n        function openPhatamPage(ev) {\n            if (ev && ev.preventDefault) ev.preventDefault();\n            const info = document.getElementById(\"voice-info\");\n            if (acWhere()) { if (info) { info.textContent = \"Bạn đang làm bài có giám sát nên chưa mở được trang ngoài. Hãy mở sau khi nộp bài.\"; info.classList.add(\"warn\"); } return false; }\n            const u = phatamUrl();\n            if (!u) { if (info) { info.textContent = \"Chưa có địa chỉ trang. Hãy mở từ link của giáo viên.\"; info.classList.add(\"warn\"); } return false; }\n            window.open(u, \"_blank\", \"noopener\");\n            return false;\n        }\n        function openPhonetics(kind, text, ipa) {\n            const t = kind === \"card\" ? phCardText() : String(text || \"\");\n            const locked = !!acWhere();\n            const rows = phAnalyse(t, ipa);\n            const a = (href, label) => locked || !href ? '<span class=\"ph-a off\">' + label + '</span>' : '<a class=\"ph-a\" href=\"' + escHtml(href) + '\" target=\"_blank\" rel=\"noopener noreferrer\">' + label + '</a>';\n            let html = \"\";\n            if (locked) html += '<div class=\"ph-note\">🛡 Bạn đang làm bài có giám sát nên chưa mở được trang ngoài (mở sẽ tính là rời trang). Hãy xem sau khi nộp bài; phần mô tả bên dưới vẫn đọc được.</div>';\n            if (!rows.length) html += '<p class=\"ph-tip\">Không thấy âm nào đặc biệt khó với người Việt trong phần này. Bạn vẫn có thể xem cách đọc từng âm trên bảng âm của Sayshape.</p>';\n            rows.forEach(r => {\n                const inf = PH_INFO[r.id];\n                const links = inf.links.map(sym => a(ssUrl(sym), \"▶ Xem cách đọc /\" + (sym === \"g\" ? \"ɡ\" : sym) + \"/\")).join(\"\");\n                html += '<div class=\"ph-row\"><div class=\"ph-sym\">' + escHtml(inf.sym) + '</div><div><div class=\"ph-title\">' + escHtml(inf.title) + '</div>' +\n                    '<div class=\"ph-tip\"><b>Trong câu này:</b> ' + r.words.map(escHtml).join(\", \") + '<br>' + escHtml(inf.tip) + '</div><div class=\"ph-links\">' + links + '</div></div></div>';\n            });\n            document.getElementById(\"ph-title\").textContent = \"👄 Cách phát âm\";\n            document.getElementById(\"ph-sub\").textContent = t.length > 140 ? t.slice(0, 137) + \"…\" : t;\n            document.getElementById(\"ph-body\").innerHTML = html;\n            const homeUrl = SS_BASE + \"?accent=\" + (voicePref.acc === \"uk\" ? \"uk\" : \"us\") + \"&lang=vi&l1=vi\";\n            const pu = phatamUrl();\n            document.getElementById(\"ph-foot\").innerHTML = a(homeUrl, \"🗺 Mở bảng 44 âm (Sayshape)\") + (pu ? a(pu, \"📚 Âm khó với người Việt\") : \"\");\n            document.getElementById(\"ph-modal\").classList.remove(\"hidden\");\n        }\n        function closePhonetics() { document.getElementById(\"ph-modal\").classList.add(\"hidden\"); }\n        document.addEventListener(\"keydown\", ev => { if (ev.key === \"Escape\") closePhonetics(); });\n\n        function speakCurrentSentence(withAnswer = false, slow = false) {\n            const q = getCurrentQuestion();\n            if (!q) return;\n            const record = appState.answers[q.id];\n            const t = qType(q);\n            if (t === \"viet_lai\" || (t === \"trac_nghiem\" && !hasBlank(q))) { speakText(plain(q.sentence), slow ? { rate: 0.55, once: true } : null); return; }\n            const reveal = withAnswer || (record && record.status);\n            const full = reveal ? completedText(q) : plain(q.sentence).replace(BLANK, \"blank\");\n            speakText(full, slow ? { rate: 0.55, once: true } : null);\n        }\n\n        // ==========================================================================\n        // 4. HÀM KIỂM TRA ĐÁP ÁN (HỖ TRỢ ANH-ANH / ANH-MỸ)\n        // ==========================================================================\n\n        // ==========================================================================\n        // v5: 4 DẠNG BÀI (Điền từ / Trắc nghiệm / Viết lại câu / Word entry)\n        // ==========================================================================\n        const TYPE_LABEL = { dien_tu: \"Điền từ\", trac_nghiem: \"Trắc nghiệm\", viet_lai: \"Viết lại câu\", word_entry: \"Word entry\" };\n        const TYPE_ORDER = [\"dien_tu\", \"trac_nghiem\", \"viet_lai\", \"word_entry\"];\n        const OPT_LETTERS = [\"A\", \"B\", \"C\", \"D\"];\n        const BLANK = \"__________\";\n        let fcSel = \"\";   // lựa chọn trắc nghiệm đang chọn (chưa nộp)\n\n        function qType(q) { return q && TYPE_LABEL[q.type] ? q.type : \"dien_tu\"; }\n        function letterOf(v) {\n            const m = String(v == null ? \"\" : v).trim().toUpperCase().match(/^\\(?([A-D])[\\).:]?$/);\n            return m ? m[1] : \"\";\n        }\n        function normAns(s) { return String(s == null ? \"\" : s).replace(/\\s+/g, \" \").trim().toLowerCase(); }\n        function normLoose(s) {\n            return String(s == null ? \"\" : s).normalize(\"NFC\").toLowerCase()\n                .replace(/[’‘`´]/g, \"'\")\n                .replace(/[^a-z0-9À-ɏ\\s']/g, \" \")\n                .replace(/\\s+/g, \" \").trim();\n        }\n        function rich(s) { return escHtml(s).replace(/\\*\\*(.+?)\\*\\*/g, '<b class=\"hl\">$1</b>'); }\n        function plain(s) { return String(s == null ? \"\" : s).replace(/\\*\\*/g, \"\"); }\n        function optText(q, letter) {\n            const i = OPT_LETTERS.indexOf(letterOf(letter));\n            return i >= 0 && q.opts ? (q.opts[i] || \"\") : \"\";\n        }\n        function displayAnswer(q, forSentence) {\n            if (qType(q) === \"trac_nghiem\") {\n                const L = letterOf(q.answer);\n                return L + \". \" + plain(optText(q, L));\n            }\n            return q.altAnswers && q.altAnswers.length ? (forSentence ? q.answer + \" (hoặc \" + q.altAnswers[0] + \")\" : q.answer + \" / \" + q.altAnswers[0]) : q.answer;\n        }\n        function hasBlank(q) { return String(q.sentence).indexOf(BLANK) >= 0; }\n        // câu hoàn chỉnh dạng chữ thuần (để đọc và hiển thị)\n        function completedText(q) {\n            const t = qType(q);\n            if (t === \"viet_lai\") return q.answer;\n            if (t === \"trac_nghiem\") return hasBlank(q) ? plain(q.sentence).replace(BLANK, plain(optText(q, q.answer))) : displayAnswer(q);\n            return plain(q.sentence).replace(BLANK, q.answer);\n        }\n        function entryOf(q) { return entriesData.find(e => e.id === q.entryId) || null; }\n        function entryHtml(e, open, collapsible) {\n            if (!e) return \"\";\n            const defs = (e.defs || []).map(d => \"<li>\" + rich(d.replace(/^\\s*\\d+[\\.\\)]\\s*/, \"\")) + \"</li>\").join(\"\");\n            const exs = (e.examples || []).map(x => \"<li>\" + rich(x) + \"</li>\").join(\"\");\n            const body = '<div class=\"entry-head\"><span class=\"entry-word\">' + escHtml(e.name) + '</span>' +\n                (e.ipa ? '<span class=\"entry-ipa\">' + escHtml(e.ipa) + '</span>' : '') +\n                (e.pos ? '<span class=\"entry-pos\">' + escHtml(e.pos) + '</span>' : '') +\n                '<button type=\"button\" class=\"btn btn-ghost\" style=\"padding:2px 9px; font-size:0.74rem;\" onclick=\"speakText(this.getAttribute(\\'data-say\\'))\" data-say=\"' + escHtml(e.name) + '\">🔊</button>' +\n                '<button type=\"button\" class=\"btn btn-ghost\" style=\"padding:2px 9px; font-size:0.74rem;\" onclick=\"openPhonetics(\\'word\\', this.getAttribute(\\'data-say\\'), this.getAttribute(\\'data-ipa\\'))\" data-say=\"' + escHtml(e.name) + '\" data-ipa=\"' + escHtml(e.ipa || \"\") + '\" title=\"Cách phát âm\">👄</button></div>' +\n                (defs ? \"<ol>\" + defs + \"</ol>\" : \"\") + (exs ? \"<ul>\" + exs + \"</ul>\" : \"\");\n            if (!collapsible) return '<div class=\"entry-card\">' + body + '</div>';\n            return '<div class=\"entry-card\"><details' + (open ? \" open\" : \"\") + '><summary>📖 Mục từ điển: ' + escHtml(e.name) + '</summary>' + body + '</details></div>';\n        }\n        function exampleWithAnswer(e, q) {\n            if (!e) return \"\";\n            const keys = [q.answer].concat(q.altAnswers || []).map(normLoose).filter(Boolean);\n            for (const ex of (e.examples || [])) {\n                const bolds = (ex.match(/\\*\\*(.+?)\\*\\*/g) || []).map(b => normLoose(b.replace(/\\*\\*/g, \"\")));\n                if (keys.some(k => bolds.indexOf(k) >= 0)) return ex;\n            }\n            for (const ex of (e.examples || [])) {\n                const n = normLoose(ex);\n                if (keys.some(k => (\" \" + n + \" \").indexOf(\" \" + k + \" \") >= 0)) return ex;\n            }\n            return \"\";\n        }\n        function unitTag(q) {\n            const bits = [];\n            if (q.unit) bits.push(splitUnit(q.unit)[0]);\n            if (q.topic) bits.push(q.topic);\n            return bits.length ? '<span class=\"q-unit-tag\">' + bits.map(escHtml).join(\" · \") + '</span>' : \"\";\n        }\n        function typeTag(q) { const t = qType(q); return '<span class=\"type-tag t-' + t + '\">' + TYPE_LABEL[t] + '</span>'; }\n        function typesPresent() { return TYPE_ORDER.filter(t => questionsData.some(q => qType(q) === t)); }\n        function fillTypeSelect(id) {\n            const sel = el(id); if (!sel) return;\n            const inScope = questionsData.filter(scopeMatch);\n            const present = TYPE_ORDER.filter(t => inScope.some(q => qType(q) === t));\n            const keep = sel.value;\n            if (typesPresent().length < 2) { sel.classList.add(\"hidden\"); sel.innerHTML = \"\"; return; }\n            sel.classList.remove(\"hidden\");\n            sel.innerHTML = '<option value=\"all\">Tất cả dạng bài</option>' + present.map(t =>\n                '<option value=\"' + t + '\">' + TYPE_LABEL[t] + \" (\" + inScope.filter(q => qType(q) === t).length + \")</option>\").join(\"\");\n            sel.value = present.indexOf(keep) >= 0 || keep === \"all\" ? (keep || \"all\") : \"all\";\n        }\n        function typeFilterValue(id) { const e = el(id); return e && !e.classList.contains(\"hidden\") ? e.value : \"all\"; }\n        function poolFor(id) {\n            const tf = typeFilterValue(id);\n            return questionsData.filter(q => (tf === \"all\" || qType(q) === tf) && scopeMatch(q));\n        }\n\n        // ==========================================================================\n        // v6. LỌC THEO UNIT / CHỦ ĐỀ / TỪ KHÓA (dùng chung cho 3 chế độ)\n        // ==========================================================================\n        const NO_UNIT = \"Chưa phân Unit\", NO_TOPIC = \"Chưa phân chủ đề\";\n        function foldS(s) { return String(s == null ? \"\" : s).normalize(\"NFD\").replace(/[\\u0300-\\u036f]/g, \"\").replace(/đ/g, \"d\").replace(/Đ/g, \"d\").toLowerCase(); }\n        function unitOf(q) { return q.unit || NO_UNIT; }\n        function topicOf(q) { return q.topic || NO_TOPIC; }\n        function scopeActive() { const s = appState.scope; return !!(s.unit || s.topic || s.q); }\n        // tìm theo câu hỏi, từ gốc, dịch nghĩa, Unit, chủ đề (KHÔNG tìm theo đáp án để khỏi lộ đáp án)\n        function searchText(q) {\n            if (q._st === undefined) q._st = foldS([q.sentence, q.root, q.translation, q.unit, q.topic, q.id].join(\" \"));\n            return q._st;\n        }\n        function scopeMatch(q) {\n            const s = appState.scope;\n            if (s.unit && unitOf(q) !== s.unit) return false;\n            if (s.topic && topicOf(q) !== s.topic) return false;\n            if (s.q) {\n                const hay = searchText(q);\n                if (!foldS(s.q).split(/\\s+/).filter(Boolean).every(w => hay.indexOf(w) >= 0)) return false;\n            }\n            return true;\n        }\n        function scopedIds() { return questionsData.filter(scopeMatch).map(q => q.id); }\n        function unitKey(u) {\n            if (u === NO_UNIT) return [3, 0, u];\n            const m = /^unit\\s*(\\d+)/i.exec(u);\n            return m ? [0, parseInt(m[1], 10), u] : [1, 0, u];\n        }\n        function cmpUnit(a, b) {\n            const x = unitKey(a), y = unitKey(b);\n            return (x[0] - y[0]) || (x[1] - y[1]) || x[2].localeCompare(y[2], \"vi\");\n        }\n        function splitUnit(u) {\n            const m = /^(unit\\s*\\d+)\\s*[:\\-–]\\s*(.+)$/i.exec(u);\n            return m ? [m[1], m[2]] : [u, \"\"];\n        }\n        function statOf(list) {\n            let done = 0, ok = 0;\n            list.forEach(q => { const r = appState.answers[q.id]; if (r && r.status) { done++; if (r.status === \"correct\") ok++; } });\n            return { total: list.length, done, ok };\n        }\n        function hasScopeData() { return questionsData.some(q => q.unit || q.topic) || questionsData.length > 30; }\n\n        function renderScopeBar() {\n            const bar = el(\"scope-bar\"); if (!bar) return;\n            const show = questionsData.length > 0 && hasScopeData();\n            bar.classList.toggle(\"hidden\", !show);\n            if (!show) return;\n            const s = appState.scope;\n            // --- thẻ Unit (đếm trên toàn bộ câu, không phụ thuộc Unit đang chọn; có tính chủ đề & từ khóa đang chọn)\n            const units = {};\n            questionsData.forEach(q => { (units[unitOf(q)] = units[unitOf(q)] || []).push(q); });\n            const names = Object.keys(units).sort(cmpUnit);\n            const showUnits = names.length > 1 || (names.length === 1 && names[0] !== NO_UNIT);\n            const all = statOf(questionsData);\n            const card = (key, no, nm, st) => {\n                const pct = st.total ? Math.round(st.done / st.total * 100) : 0;\n                const acc = st.done ? Math.round(st.ok / st.done * 100) : 0;\n                return '<button class=\"unit-card' + (s.unit === key ? \" active\" : \"\") + '\" data-u=\"' + escHtml(key) + '\" onclick=\"pickUnit(this.dataset.u)\">' +\n                    '<span class=\"u-no\">' + escHtml(no) + '</span><span class=\"u-name\">' + escHtml(nm) + '</span>' +\n                    '<span class=\"u-bar\"><i class=\"' + (st.done && st.done === st.total ? \"ok\" : \"\") + '\" style=\"width:' + pct + '%\"></i></span>' +\n                    '<span class=\"u-meta\"><span>' + st.total + ' câu</span><span>' + (st.done ? \"✔ \" + st.done + \"/\" + st.total + \" · \" + acc + \"% đúng\" : \"chưa làm\") + '</span></span></button>';\n            };\n            el(\"unit-cards\").classList.toggle(\"hidden\", !showUnits);\n            el(\"unit-cards\").innerHTML = showUnits\n                ? card(\"\", \"Tất cả\", \"Mọi Unit\", all) + names.map(u => { const p = splitUnit(u); return card(u, p[0], p[1], statOf(units[u])); }).join(\"\")\n                : \"\";\n            // --- chip chủ đề (theo Unit đang chọn)\n            const inUnit = questionsData.filter(q => !s.unit || unitOf(q) === s.unit);\n            const topics = {};\n            inUnit.forEach(q => { (topics[topicOf(q)] = topics[topicOf(q)] || []).push(q); });\n            const tnames = Object.keys(topics).sort((a, b) => (a === NO_TOPIC) - (b === NO_TOPIC) || a.localeCompare(b, \"vi\"));\n            const showTopics = tnames.length > 1 || (tnames.length === 1 && tnames[0] !== NO_TOPIC);\n            el(\"topic-chips\").classList.toggle(\"hidden\", !showTopics);\n            el(\"topic-chips\").innerHTML = showTopics\n                ? '<button class=\"t-chip' + (!s.topic ? \" active\" : \"\") + '\" data-t=\"\" onclick=\"pickTopic(this.dataset.t)\">Tất cả chủ đề<small>' + inUnit.length + '</small></button>' +\n                  tnames.map(tn => '<button class=\"t-chip' + (s.topic === tn ? \" active\" : \"\") + '\" data-t=\"' + escHtml(tn) + '\" onclick=\"pickTopic(this.dataset.t)\">' + escHtml(tn) + '<small>' + topics[tn].length + '</small></button>').join(\"\")\n                : \"\";\n            // --- tóm tắt\n            const ids = questionsData.filter(scopeMatch), st = statOf(ids);\n            const sum = el(\"scope-summary\");\n            if (scopeActive()) {\n                const parts = [];\n                if (s.unit) parts.push(s.unit);\n                if (s.topic) parts.push(s.topic);\n                if (s.q) parts.push('\"' + s.q + '\"');\n                sum.innerHTML = \"Đang chọn: <b>\" + escHtml(parts.join(\" › \")) + \"</b> — \" + ids.length + \" / \" + questionsData.length + \" câu\" +\n                    (ids.length ? \" · đã làm \" + st.done + (st.done ? \", đúng \" + st.ok : \"\") : \"\");\n            } else {\n                sum.innerHTML = \"Đang ôn <b>tất cả \" + questionsData.length + \" câu</b>. Chọn Unit hoặc chủ đề để học từng phần.\";\n            }\n            el(\"scope-clear\").classList.toggle(\"hidden\", !scopeActive());\n        }\n\n        // áp dụng bộ lọc: dựng lại danh sách câu ở cả 3 chế độ\n        function applyScope(keepWorksheet) {\n            renderScopeBar();\n            fillTypeSelect(\"fc-type-filter\"); fillTypeSelect(\"ws-type-filter\");\n            setFlashcardFilter(\"all\", true);\n            if (!keepWorksheet) {\n                if (appState.currentMode === \"worksheet\") generateWorksheet(appState.wsQuizSize);\n                else el(\"worksheet-list\").innerHTML = \"\";\n            }\n            if (appState.currentMode === \"handbook\") renderWordFamilies();\n            const allBtn = el(\"ws-size-all\"); if (allBtn) allBtn.innerText = \"📋 Tất cả \" + Math.min(poolFor(\"ws-type-filter\").length, WS_ALL_MAX) + \" câu\";\n        }\n        function pickUnit(u) {\n            const s = appState.scope;\n            s.unit = (s.unit === u) ? \"\" : u;\n            s.topic = \"\";\n            applyScope();\n        }\n        function pickTopic(tn) {\n            const s = appState.scope;\n            s.topic = (s.topic === tn) ? \"\" : tn;\n            applyScope();\n        }\n        let scopeTimer = null;\n        function onScopeSearch() {\n            clearTimeout(scopeTimer);\n            scopeTimer = setTimeout(() => { appState.scope.q = el(\"scope-search\").value.trim(); applyScope(); }, 220);\n        }\n        function clearScope(silent) {\n            appState.scope = { unit: \"\", topic: \"\", q: \"\" };\n            const si = el(\"scope-search\"); if (si) si.value = \"\";\n            if (!silent) applyScope(); else renderScopeBar();\n        }\n\n        function checkAnswerMatch(userInput, questionObj) {\n            const t = qType(questionObj);\n            const raw = userInput == null ? \"\" : String(userInput);\n            if (t === \"trac_nghiem\") {\n                const L = letterOf(raw);\n                return !!L && L === letterOf(questionObj.answer);\n            }\n            const norm = (t === \"viet_lai\" || t === \"word_entry\") ? normLoose : normAns;\n            const clean = norm(raw);\n            if (!clean) return false;\n            if (clean === norm(questionObj.answer)) return true;\n            return !!(questionObj.altAnswers && questionObj.altAnswers.some(a => norm(a) === clean));\n        }\n\n        let _qIdx = null, _qIdxFor = null;\n        function getQuestionById(id) {\n            if (_qIdxFor !== questionsData || !_qIdx || _qIdx.size !== questionsData.length) {\n                _qIdx = new Map(); questionsData.forEach(q => _qIdx.set(q.id, q)); _qIdxFor = questionsData;\n            }\n            return _qIdx.get(id);\n        }\n\n        function getCurrentQuestion() {\n            const qId = appState.currentOrder[appState.currentIndex];\n            return getQuestionById(qId) || questionsData[0];\n        }\n\n        // ==========================================================================\n        // 5. CẬP NHẬT THỐNG KÊ TIẾN ĐỘ (STATS BAR)\n        // ==========================================================================\n        function updateStatsUI() {\n            try { if (typeof renderScopeBar === \"function\" && questionsData.length) renderScopeBar(); } catch (e) {}\n            let doneCount = 0;\n            let correctCount = 0;\n            let wrongCount = 0;\n\n            questionsData.forEach(q => {\n                const rec = appState.answers[q.id];\n                if (rec && rec.status) {\n                    doneCount++;\n                    if (rec.status === \"correct\") correctCount++;\n                    else wrongCount++;\n                }\n            });\n\n            document.getElementById(\"stat-progress\").innerText = `${doneCount} / ${questionsData.length}`;\n            document.getElementById(\"stat-correct\").innerText = correctCount;\n            document.getElementById(\"stat-wrong\").innerText = wrongCount;\n            document.getElementById(\"stat-streak\").innerText = appState.streak;\n\n            const pct = Math.round((doneCount / questionsData.length) * 100);\n            document.getElementById(\"main-progress-bar\").style.width = `${pct}%`;\n        }\n\n        // ==========================================================================\n        // 6. CHẾ ĐỘ 1: LUYỆN TẬP TỪNG CÂU (SINGLE QUESTION PRACTICE)\n        // ==========================================================================\n        function setFlashcardFilter(filterType, fromSelect) {\n            const pool = poolFor(\"fc-type-filter\");\n            const pick = f => {\n                if (f === \"all\") return pool.map(q => q.id);\n                if (f === \"unanswered\") return pool.filter(q => !appState.answers[q.id] || !appState.answers[q.id].status).map(q => q.id);\n                return pool.filter(q => appState.answers[q.id] && appState.answers[q.id].status && appState.answers[q.id].status !== \"correct\").map(q => q.id);\n            };\n            let newOrder = pick(filterType);\n\n            if (newOrder.length === 0) {\n                if (fromSelect) { filterType = \"all\"; newOrder = pick(\"all\"); }\n                else {\n                    const btn = document.getElementById(`fc-filter-${filterType}`);\n                    const oldText = btn.innerText;\n                    btn.innerText = \"🎉 Trống!\";\n                    setTimeout(() => { btn.innerText = oldText; }, 1400);\n                    return;\n                }\n            }\n\n            appState.fcFilter = filterType;\n            [\"all\", \"unanswered\", \"wrong\"].forEach(f => {\n                document.getElementById(`fc-filter-${f}`).className = (f === filterType) ? \"btn btn-accent\" : \"btn btn-ghost\";\n            });\n            el(\"fc-filter-all\").innerText = \"Tất cả (\" + pool.length + \")\";\n\n            appState.currentOrder = newOrder;\n            appState.currentIndex = 0;\n            renderFlashcardView();\n        }\n\n        function onFcTypeChange() { setFlashcardFilter(appState.fcFilter, true); }\n\n        function resetFcFilterUI() {\n            const sel = el(\"fc-type-filter\"); if (sel) sel.value = \"all\";\n            appState.fcFilter = \"all\";\n            [\"all\", \"unanswered\", \"wrong\"].forEach(f => { el(\"fc-filter-\" + f).className = (f === \"all\") ? \"btn btn-accent\" : \"btn btn-ghost\"; });\n            el(\"fc-filter-all\").innerText = \"Tất cả (\" + poolFor(\"fc-type-filter\").length + \")\";\n        }\n\n        function selectFcOption(L) {\n            const q = getCurrentQuestion();\n            if (!q || qType(q) !== \"trac_nghiem\") return;\n            fcSel = L;\n            el(\"fc-options\").querySelectorAll(\".opt-btn\").forEach(b => {\n                b.classList.remove(\"opt-correct\", \"opt-wrong\");\n                b.classList.toggle(\"selected\", b.getAttribute(\"data-l\") === L);\n            });\n        }\n\n        const PILL_PAGE = 100;\n        function pillGo(dir) {\n            const pages = Math.ceil(appState.currentOrder.length / PILL_PAGE);\n            const cur = typeof appState.pillPage === \"number\" ? appState.pillPage : Math.floor(appState.currentIndex / PILL_PAGE);\n            appState.pillPage = Math.max(0, Math.min(pages - 1, cur + dir));\n            renderFlashcardPills(true);\n        }\n        function renderFlashcardPills(keepPage) {\n            const container = document.getElementById(\"fc-pills-container\");\n            container.innerHTML = \"\";\n            const n = appState.currentOrder.length;\n            let from = 0, to = n;\n            if (n > PILL_PAGE * 1.5) {\n                const cp = Math.floor(appState.currentIndex / PILL_PAGE);\n                if (!keepPage || typeof appState.pillPage !== \"number\") appState.pillPage = cp;\n                from = appState.pillPage * PILL_PAGE; to = Math.min(n, from + PILL_PAGE);\n                const prev = document.createElement(\"button\");\n                prev.className = \"q-pill\"; prev.innerText = \"‹\"; prev.title = \"100 câu trước\"; prev.disabled = appState.pillPage === 0;\n                prev.onclick = () => pillGo(-1); container.appendChild(prev);\n            } else { appState.pillPage = undefined; }\n\n            for (let idx = from; idx < to; idx++) {\n                const qId = appState.currentOrder[idx];\n                const btn = document.createElement(\"button\");\n                btn.className = \"q-pill\";\n                btn.innerText = qId;\n\n                const rec = appState.answers[qId];\n                if (rec && rec.status === \"correct\") btn.classList.add(\"done-correct\");\n                else if (rec && (rec.status === \"wrong\" || rec.status === \"revealed\")) btn.classList.add(\"done-wrong\");\n\n                if (idx === appState.currentIndex) btn.classList.add(\"current\");\n\n                btn.onclick = () => {\n                    appState.currentIndex = idx;\n                    saveState();\n                    renderFlashcardView();\n                };\n                container.appendChild(btn);\n            }\n            if (n > PILL_PAGE * 1.5) {\n                const lab = document.createElement(\"span\");\n                lab.className = \"muted\"; lab.style.cssText = \"font-size:0.75rem;align-self:center;padding:0 4px;\";\n                lab.innerText = (from + 1) + \"–\" + to + \" / \" + n;\n                container.appendChild(lab);\n                const next = document.createElement(\"button\");\n                next.className = \"q-pill\"; next.innerText = \"›\"; next.title = \"100 câu sau\"; next.disabled = to >= n;\n                next.onclick = () => pillGo(1); container.appendChild(next);\n            }\n        }\n\n        function renderFlashcardView() {\n            renderFlashcardPills();\n            const emptyScope = appState.currentOrder.length === 0;\n            el(\"fc-empty\").classList.toggle(\"hidden\", !emptyScope);\n            el(\"fc-card\").classList.toggle(\"hidden\", emptyScope);\n            el(\"fc-nav-row\").classList.toggle(\"hidden\", emptyScope);\n            if (emptyScope) return;\n            const q = getCurrentQuestion();\n            if (!q) return;\n\n            const t = qType(q);\n            const rec = appState.answers[q.id];\n            const wf = t === \"dien_tu\" ? wordFamilies.find(f => f.id === q.familyId) : null;\n            const answered = !!(rec && rec.status);\n            const isOk = answered && rec.status === \"correct\";\n            fcSel = \"\";\n\n            document.getElementById(\"fc-q-number\").innerHTML = (scopeActive() ? `CÂU ${q.id} (${appState.currentIndex + 1} / ${appState.currentOrder.length})` : `CÂU ${q.id} / ${questionsData.length}`) + (typesPresent().length > 1 ? typeTag(q) : \"\") + unitTag(q);\n            el(\"fc-root-word\").innerText = t === \"dien_tu\" ? \"(\" + q.root + \")\" : (t === \"word_entry\" ? \"Word entry\" : TYPE_LABEL[t]);\n\n            // gia đình từ chỉ có ở dạng Điền từ\n            const hintFamEl = document.getElementById(\"fc-hint-family\");\n            hintFamEl.classList.add(\"hidden\");\n            hintFamEl.innerHTML = wf ? `<strong>💡 Gia đình từ (${wf.root}):</strong> ${wf.content}` : \"\";\n            el(\"fc-fam-btn\").classList.toggle(\"hidden\", t !== \"dien_tu\");\n\n            // mục từ điển (Word entry)\n            el(\"fc-entry\").innerHTML = t === \"word_entry\" ? entryHtml(entryOf(q), true, false) : \"\";\n\n            const inputEl = el(\"fc-input\"), areaEl = el(\"fc-textarea\"), optEl = el(\"fc-options\");\n            const expBox = document.getElementById(\"fc-explanation-box\");\n            inputEl.classList.toggle(\"hidden\", t === \"trac_nghiem\" || t === \"viet_lai\");\n            areaEl.classList.toggle(\"hidden\", t !== \"viet_lai\");\n            optEl.classList.toggle(\"hidden\", t !== \"trac_nghiem\");\n            inputEl.placeholder = t === \"word_entry\" ? \"Gõ cụm 2–3 từ lấy từ mục từ điển...\" : \"Gõ dạng từ đúng vào đây...\";\n            const box = t === \"viet_lai\" ? areaEl : inputEl;\n\n            // phần câu hỏi\n            const fillClass = isOk ? \"blank-filled-correct\" : \"blank-filled-wrong\";\n            const filled = displayAnswer(q, false);\n            let sHtml;\n            if (t === \"viet_lai\") {\n                sHtml = rich(q.sentence) + (q.root ? `<div class=\"rw-hint\">${rich(q.root)}</div>` : \"\");\n            } else if (t === \"trac_nghiem\") {\n                const fill = answered ? `<span class=\"${fillClass}\">${rich(optText(q, q.answer))}</span>` : `<span class=\"blank-highlight\">${BLANK}</span>`;\n                sHtml = hasBlank(q) ? rich(q.sentence).replace(BLANK, fill) : rich(q.sentence);\n            } else {\n                const fill = answered ? `<span class=\"${fillClass}\">${escHtml(filled)}</span>` : `<span class=\"blank-highlight\">${BLANK}</span>`;\n                sHtml = rich(q.sentence).replace(BLANK, fill) + (t === \"dien_tu\" ? ` <span class=\"root-inline\">(${escHtml(q.root)})</span>` : \"\");\n            }\n            document.getElementById(\"fc-sentence\").innerHTML = sHtml;\n\n            // lựa chọn trắc nghiệm\n            if (t === \"trac_nghiem\") {\n                const correctL = letterOf(q.answer), chosenL = answered ? letterOf(rec.userText) : \"\";\n                optEl.innerHTML = OPT_LETTERS.map((L, i) => {\n                    const txt = (q.opts || [])[i];\n                    if (!txt) return \"\";\n                    let cls = \"opt-btn\";\n                    if (answered) {\n                        if (L === correctL) cls += \" opt-correct\";\n                        else if (L === chosenL && !isOk) cls += \" opt-wrong\";\n                    }\n                    return `<button type=\"button\" class=\"${cls}\" data-l=\"${L}\" onclick=\"selectFcOption('${L}')\"><span class=\"opt-letter\">${L}</span><span>${rich(txt)}</span></button>`;\n                }).join(\"\");\n            } else optEl.innerHTML = \"\";\n\n            if (answered) {\n                if (t !== \"trac_nghiem\") {\n                    box.value = rec.userText || q.answer;\n                    box.className = (t === \"viet_lai\" ? \"answer-textarea \" : \"answer-input \") + (isOk ? \"input-correct\" : \"input-wrong\");\n                }\n                showFlashcardExplanation(q, wf, rec);\n            } else {\n                box.value = \"\";\n                box.className = t === \"viet_lai\" ? \"answer-textarea\" : \"answer-input\";\n                expBox.classList.add(\"hidden\");\n                if (t !== \"trac_nghiem\") setTimeout(() => box.focus(), 60);\n            }\n            inputEl.classList.toggle(\"hidden\", t === \"trac_nghiem\" || t === \"viet_lai\");\n            areaEl.classList.toggle(\"hidden\", t !== \"viet_lai\");\n        }\n\n        function showFlashcardExplanation(q, wf, rec) {\n            const t = qType(q);\n            const expBox = document.getElementById(\"fc-explanation-box\");\n            const statusEl = document.getElementById(\"fc-exp-status\");\n            expBox.classList.remove(\"hidden\");\n\n            const displayAns = t === \"trac_nghiem\" ? escHtml(displayAnswer(q)) : escHtml(displayAnswer(q, true));\n            const mine = t === \"trac_nghiem\" ? `Bạn chọn: ${escHtml(letterOf(rec.userText))}` : `Bạn gõ: \"${escHtml(rec.userText)}\"`;\n\n            if (rec.status === \"correct\") {\n                expBox.className = \"explanation-box correct\";\n                statusEl.innerHTML = `<span style=\"color:#4ade80;\">✅ Chính xác! Đáp án: <u style=\"color:#fef08a;\">${displayAns}</u></span>`;\n            } else if (rec.status === \"revealed\") {\n                expBox.className = \"explanation-box wrong\";\n                statusEl.innerHTML = `<span style=\"color:#fbbf24;\">👀 Đáp án chuẩn: <u style=\"color:#fff;\">${displayAns}</u></span>`;\n            } else {\n                expBox.className = \"explanation-box wrong\";\n                statusEl.innerHTML = `<span style=\"color:#fda4af;\">❌ Chưa đúng (${mine}) &rarr; Đáp án đúng: <u style=\"color:#fef08a;\">${displayAns}</u></span>`;\n            }\n\n            el(\"fc-exp-full-label\").innerText = t === \"viet_lai\" ? \"✅ Câu viết lại đúng:\" : ((t === \"trac_nghiem\" && !hasBlank(q)) ? \"✅ Đáp án:\" : \"✅ Câu hoàn chỉnh:\");\n            document.getElementById(\"fc-exp-full-sentence\").innerText = \" \" + (t === \"dien_tu\" ? plain(q.sentence).replace(BLANK, `[${q.answer}]`) : completedText(q));\n            document.getElementById(\"fc-exp-grammar\").innerText = \" \" + (q.explain || \"\");\n            document.getElementById(\"fc-exp-translation\").innerText = \" \" + (q.translation || \"\");\n            document.getElementById(\"fc-exp-family\").innerText = wf ? \" \" + wf.content : \"\";\n            el(\"fc-exp-family-wrap\").style.display = wf ? \"\" : \"none\";\n            const ex = t === \"word_entry\" ? exampleWithAnswer(entryOf(q), q) : \"\";\n            el(\"fc-exp-entry-wrap\").style.display = ex ? \"\" : \"none\";\n            el(\"fc-exp-entry\").innerHTML = ex ? \" \" + rich(ex) : \"\";\n        }\n\n        function checkFlashcardAnswer() {\n            const q = getCurrentQuestion();\n            const t = qType(q);\n            const rec0 = appState.answers[q.id];\n            if (t === \"trac_nghiem\" && rec0 && rec0.status && !fcSel) return;\n            const inputEl = t === \"viet_lai\" ? el(\"fc-textarea\") : el(\"fc-input\");\n            const val = t === \"trac_nghiem\" ? fcSel : inputEl.value.trim();\n\n            if (!val) {\n                if (t === \"trac_nghiem\") {\n                    const b = el(\"fc-submit-btn\"), old = b.innerText;\n                    b.innerText = \"Hãy chọn A, B, C hoặc D\";\n                    setTimeout(() => { b.innerText = old; }, 1400);\n                } else inputEl.focus();\n                return;\n            }\n\n            const isCorrect = checkAnswerMatch(val, q);\n            playFeedbackTone(isCorrect);\n\n            appState.streak = isCorrect ? appState.streak + 1 : 0;\n            appState.answers[q.id] = {\n                userText: val,\n                status: isCorrect ? \"correct\" : \"wrong\"\n            };\n\n            saveState();\n            updateStatsUI();\n            renderFlashcardView();\n        }\n\n        function handleFlashcardEnter() {\n            const q = getCurrentQuestion();\n            const t = qType(q);\n            const rec = appState.answers[q.id];\n            const inputEl = t === \"viet_lai\" ? el(\"fc-textarea\") : el(\"fc-input\");\n\n            if (rec && rec.status && (t === \"trac_nghiem\" ? !fcSel : inputEl.value.trim() === (rec.userText || q.answer))) {\n                nextFlashcard();\n            } else {\n                checkFlashcardAnswer();\n            }\n        }\n\n        function revealFlashcardAnswer() {\n            const q = getCurrentQuestion();\n            appState.streak = 0;\n            appState.answers[q.id] = {\n                userText: q.answer,\n                status: \"revealed\"\n            };\n            saveState();\n            updateStatsUI();\n            renderFlashcardView();\n        }\n\n        function toggleFamilyHint() {\n            document.getElementById(\"fc-hint-family\").classList.toggle(\"hidden\");\n        }\n\n        function nextFlashcard() {\n            if (!appState.currentOrder.length) return;\n            appState.currentIndex = (appState.currentIndex + 1) % appState.currentOrder.length;\n            saveState();\n            renderFlashcardView();\n        }\n\n        function prevFlashcard() {\n            if (!appState.currentOrder.length) return;\n            appState.currentIndex = (appState.currentIndex - 1 + appState.currentOrder.length) % appState.currentOrder.length;\n            saveState();\n            renderFlashcardView();\n        }\n\n        function shuffleFlashcards() {\n            if (!appState.currentOrder.length) return;\n            appState.currentOrder.sort(() => Math.random() - 0.5);\n            appState.currentIndex = 0;\n            renderFlashcardView();\n        }\n\n        // ==========================================================================\n        // 7. CHẾ ĐỘ 2: LÀM ĐỀ TỔNG HỢP (WORKSHEET MODE: 10, 20 HOẶC 40 CÂU)\n        // ==========================================================================\n        const WS_ALL_MAX = 300;\n        function generateWorksheet(size) {\n            appState.wsQuizSize = size;\n            appState.wsSubmitted = false;\n\n            [10, 20, 1000].forEach(sz => {\n                const btn = document.getElementById(sz === 1000 ? \"ws-size-all\" : `ws-size-${sz}`);\n                if (btn) btn.className = (sz === size) ? \"btn btn-accent\" : \"btn btn-ghost\";\n            });\n\n            const pool = poolFor(\"ws-type-filter\");\n            const allIds = pool.map(q => q.id);\n            let ids;\n            if (size >= 1000 && allIds.length > WS_ALL_MAX) size = WS_ALL_MAX;   // kho lớn: \"Tất cả\" lấy tối đa 300 câu ngẫu nhiên để máy không nặng\n            if (size >= 1000 || size >= allIds.length) {\n                ids = allIds;\n            } else {\n                ids = [...allIds].sort(() => Math.random() - 0.5).slice(0, size);\n            }\n            // sắp theo dạng bài như đề thi (giữ nguyên thứ tự trong cùng dạng)\n            const rank = id => TYPE_ORDER.indexOf(qType(getQuestionById(id)));\n            ids = ids.map((id, i) => ({ id, i })).sort((a, b) => (rank(a.id) - rank(b.id)) || (a.i - b.i)).map(o => o.id);\n            appState.wsQuestionIds = ids;\n            const allBtn = el(\"ws-size-all\");\n            if (allBtn) allBtn.innerText = \"📋 Tất cả \" + Math.min(allIds.length, WS_ALL_MAX) + \" câu\";\n\n            document.getElementById(\"ws-score-banner\").classList.add(\"hidden\");\n            renderWorksheetItems();\n            acOnMode(appState.currentMode);\n        }\n\n        // Đề lớn (hàng trăm câu): dựng 30 câu đầu ngay, phần còn lại dựng dần để máy không bị đứng.\n        let wsRender = { token: 0, entryShown: null, next: 0, timer: 0 };\n        function wsFlush() {\n            if (wsRender.next < appState.wsQuestionIds.length) { clearTimeout(wsRender.timer); wsRenderRange(wsRender.token, appState.wsQuestionIds.length); }\n        }\n        function wsRenderRange(token, until) {\n            if (token !== wsRender.token) return;\n            const ids = appState.wsQuestionIds;\n            const end = Math.min(until, ids.length);\n            if (wsRender.next >= end) return;\n            wsRenderItems(wsRender.next, end);\n            wsRender.next = end;\n            if (end < ids.length) wsRender.timer = setTimeout(() => wsRenderRange(token, end + 40), 16);\n        }\n        function renderWorksheetItems() {\n            const listEl = document.getElementById(\"worksheet-list\");\n            listEl.innerHTML = \"\";\n            wsRender.token++; clearTimeout(wsRender.timer);\n            wsRender.entryShown = {}; wsRender.next = 0;\n            const first = Math.min(30, appState.wsQuestionIds.length);\n            wsRenderRange(wsRender.token, first);\n            if (wsRender.next < appState.wsQuestionIds.length) wsRender.timer = setTimeout(() => wsRenderRange(wsRender.token, wsRender.next + 40), 16);\n        }\n        function wsRenderItems(from, to) {\n            const listEl = document.getElementById(\"worksheet-list\");\n            const entryShown = wsRender.entryShown;\n            const multi = typesPresent().length > 1;\n            const idsAll = appState.wsQuestionIds;\n\n            for (let index = from; index < to; index++) {\n                const qId = idsAll[index];\n                const q = getQuestionById(qId);\n                const t = qType(q);\n                const wf = t === \"dien_tu\" ? wordFamilies.find(f => f.id === q.familyId) : null;\n                const rec = appState.answers[q.id];\n                const savedVal = (rec && !rec.blank && rec.status !== \"revealed\" && rec.userText) ? rec.userText : \"\";\n\n                const item = document.createElement(\"div\");\n                item.className = \"ws-item\";\n                item.id = `ws-card-${q.id}`;\n\n                let body = \"\";\n                if (t === \"trac_nghiem\") {\n                    const saved = letterOf(savedVal);\n                    body = `<div class=\"ws-sentence\">${rich(q.sentence)}</div><div class=\"ws-opts\" id=\"ws-opts-${q.id}\">` +\n                        OPT_LETTERS.map((L, i) => {\n                            const txt = (q.opts || [])[i];\n                            if (!txt) return \"\";\n                            return `<label class=\"opt-btn${saved === L ? \" selected\" : \"\"}\" data-l=\"${L}\"><input type=\"radio\" name=\"ws-opt-${q.id}\" value=\"${L}\" ${saved === L ? \"checked\" : \"\"} onchange=\"wsPick(${q.id})\"><span class=\"opt-letter\">${L}</span><span>${rich(txt)}</span></label>`;\n                        }).join(\"\") + `</div>`;\n                } else if (t === \"viet_lai\") {\n                    body = `<div class=\"ws-sentence\">${rich(q.sentence)}</div>` + (q.root ? `<div class=\"rw-hint\">${rich(q.root)}</div>` : \"\") +\n                        `<textarea id=\"ws-inp-${q.id}\" class=\"ws-textarea\" placeholder=\"gõ câu viết lại...\" autocomplete=\"off\" spellcheck=\"false\">${escHtml(savedVal)}</textarea>`;\n                } else {\n                    const parts = rich(q.sentence).split(BLANK);\n                    body = `<div class=\"ws-sentence\">${parts[0]}\n                        <input type=\"text\" id=\"ws-inp-${q.id}\" class=\"ws-inline-input\" style=\"${t === \"word_entry\" ? \"width:230px;\" : \"\"}\" value=\"${escHtml(savedVal)}\" placeholder=\"${t === \"word_entry\" ? \"gõ cụm 2–3 từ...\" : \"gõ từ...\"}\" autocomplete=\"off\" spellcheck=\"false\">\n                        ${parts[1] || \"\"}${t === \"dien_tu\" ? ` <span class=\"root-inline\">(${escHtml(q.root)})</span>` : \"\"}</div>`;\n                }\n\n                let entryBlock = \"\";\n                if (t === \"word_entry\") {\n                    const e = entryOf(q);\n                    if (e) { entryBlock = entryHtml(e, !entryShown[e.id], true); entryShown[e.id] = 1; }\n                }\n\n                const sayText = (t === \"viet_lai\" || (t === \"trac_nghiem\" && !hasBlank(q))) ? plain(q.sentence) : completedText(q);\n\n                item.innerHTML = `\n                    <div class=\"ws-top-line\">\n                        <span class=\"ws-q-num\">Câu ${index + 1} <span style=\"color:var(--text-muted); font-weight:600;\">(#${q.id})</span>${multi ? typeTag(q) : \"\"}${unitTag(q)}</span>\n                        <button onclick=\"speakText(this.getAttribute('data-say'))\" data-say=\"${escHtml(sayText)}\" class=\"btn btn-ghost\" style=\"padding:4px 10px; font-size:0.78rem;\">\n                            🔊 Nghe câu\n                        </button>\n                        <button onclick=\"openPhonetics('word', this.getAttribute('data-say'), '')\" data-say=\"${escHtml(sayText)}\" class=\"btn btn-ghost\" style=\"padding:4px 10px; font-size:0.78rem;\" title=\"Cách phát âm\">👄</button>\n                    </div>\n                    ${entryBlock}\n                    ${body}\n                    <div id=\"ws-exp-${q.id}\" class=\"ws-feedback hidden\">\n                        <div style=\"margin-bottom:5px; font-weight:800;\" id=\"ws-status-line-${q.id}\"></div>\n                        <div style=\"margin-bottom:4px;\"><strong>📌 Giải thích:</strong> ${rich(q.explain || \"\")}</div>\n                        <div style=\"margin-bottom:4px; color:#fde68a;\"><strong>🇻🇳 Dịch nghĩa:</strong> ${escHtml(q.translation || \"\")}</div>\n                        ${wf ? `<div style=\"color:#c4b5fd; font-size:0.85rem;\"><strong>🌳 Gia đình từ:</strong> ${wf.content}</div>` : \"\"}\n                    </div>\n                `;\n                listEl.appendChild(item);\n            }\n        }\n\n        function wsPick(qId) {\n            const box = el(\"ws-opts-\" + qId); if (!box) return;\n            const chosen = box.querySelector(\"input:checked\");\n            box.querySelectorAll(\".opt-btn\").forEach(l => l.classList.toggle(\"selected\", !!chosen && l.getAttribute(\"data-l\") === chosen.value));\n        }\n\n        function submitWorksheetQuiz(opts) {\n            wsFlush();\n            if (acOn() && appState.wsSubmitted && !(opts && opts.auto)) return;   // đề có giám sát: nộp một lần\n            let correctCount = 0;\n            const total = appState.wsQuestionIds.length;\n\n            appState.wsQuestionIds.forEach(qId => {\n                const q = getQuestionById(qId);\n                const t = qType(q);\n                const card = document.getElementById(`ws-card-${qId}`);\n                const expEl = document.getElementById(`ws-exp-${qId}`);\n                const statusLine = document.getElementById(`ws-status-line-${qId}`);\n\n                let userVal = \"\";\n                const inp = document.getElementById(`ws-inp-${qId}`);\n                if (t === \"trac_nghiem\") {\n                    const c = card.querySelector(`input[name=\"ws-opt-${qId}\"]:checked`);\n                    userVal = c ? c.value : \"\";\n                } else userVal = inp ? inp.value.trim() : \"\";\n                const isOk = checkAnswerMatch(userVal, q);\n                const displayAns = escHtml(displayAnswer(q, false));\n\n                if (isOk) correctCount++;\n                card.className = \"ws-item \" + (isOk ? \"ws-correct\" : \"ws-wrong\");\n                if (inp) inp.className = (t === \"viet_lai\" ? \"ws-textarea \" : \"ws-inline-input \") + (isOk ? \"ws-inp-ok\" : \"ws-inp-bad\");\n                if (t === \"trac_nghiem\") {\n                    const cl = letterOf(q.answer), ch = letterOf(userVal);\n                    card.querySelectorAll(\".opt-btn\").forEach(l => {\n                        const L = l.getAttribute(\"data-l\");\n                        l.classList.remove(\"selected\");\n                        if (L === cl) l.classList.add(\"opt-correct\");\n                        else if (L === ch && !isOk) l.classList.add(\"opt-wrong\");\n                        const r = l.querySelector(\"input\"); if (r) r.disabled = true;\n                    });\n                }\n                if (isOk) {\n                    statusLine.innerHTML = `<span style=\"color:#4ade80;\">✅ Chính xác! Đáp án: ${displayAns}</span>`;\n                } else {\n                    const userNote = userVal ? (t === \"trac_nghiem\" ? `(Bạn chọn: ${escHtml(userVal)})` : `(Bạn điền: \"${escHtml(userVal)}\")`) : \"(Chưa làm)\";\n                    statusLine.innerHTML = `<span style=\"color:#fda4af;\">❌ Chưa đúng ${userNote} &rarr; Đáp án chuẩn: <u style=\"color:#fef08a;\">${displayAns}</u></span>`;\n                }\n                if (t === \"word_entry\") {\n                    const ex = exampleWithAnswer(entryOf(q), q);\n                    if (ex) statusLine.innerHTML += `<div style=\"margin-top:4px; font-weight:600; color:#6ee7b7;\">📖 ${rich(ex)}</div>`;\n                }\n\n                expEl.classList.remove(\"hidden\");\n\n                // Đồng bộ kết quả vào tiến độ chung\n                appState.answers[qId] = {\n                    userText: userVal || q.answer,\n                    status: isOk ? \"correct\" : \"wrong\",\n                    blank: !userVal\n                };\n            });\n\n            const score10 = ((correctCount / total) * 10).toFixed(1);\n            const pct = Math.round((correctCount / total) * 100);\n\n            const banner = document.getElementById(\"ws-score-banner\");\n            banner.classList.remove(\"hidden\");\n            document.getElementById(\"ws-score-title\").innerText = `Điểm số: ${score10} / 10`;\n            document.getElementById(\"ws-score-detail\").innerText = `Số câu đúng: ${correctCount} / ${total} câu (${pct}%)`;\n\n            playFeedbackTone(correctCount >= total / 2);\n            saveState();\n            updateStatsUI();\n            acOnSubmit(opts);\n            banner.scrollIntoView({ behavior: \"smooth\", block: \"center\" });\n        }\n\n        // ==========================================================================\n        // 8. CHẾ ĐỘ 3: SỔ TAY GIA ĐÌNH TỪ (WORD FAMILIES HANDBOOK)\n        // ==========================================================================\n        function renderWordFamilies() {\n            const query = (document.getElementById(\"wf-search-input\")?.value || \"\").trim().toLowerCase();\n            const container = document.getElementById(\"wf-grid-container\");\n            container.innerHTML = \"\";\n\n            const scoped = scopeActive() ? new Set(scopedIds()) : null;\n            const inScope = q => !scoped || scoped.has(q.id);\n            const note = el(\"hb-scope-note\");\n            if (note) { note.classList.toggle(\"hidden\", !scoped); if (scoped) note.innerText = \"Đang lọc theo phần ôn tập đã chọn: chỉ hiện gia đình từ / mục từ điển có bài tập trong phần này.\"; }\n\n            const filtered = wordFamilies.filter(wf =>\n                (wf.root.toLowerCase().includes(query) || wf.content.toLowerCase().includes(query)) &&\n                (!scoped || questionsData.some(q => qType(q) === \"dien_tu\" && q.familyId === wf.id && inScope(q)))\n            );\n\n            filtered.forEach(wf => {\n                const relatedQs = questionsData.filter(q => qType(q) === \"dien_tu\" && q.familyId === wf.id);\n                const branches = wf.content.split(\"→\").map(s => s.trim());\n\n                const card = document.createElement(\"div\");\n                card.className = \"wf-card\";\n\n                const branchesHtml = branches.map((b, i) => `\n                    <div class=\"wf-branch\" style=\"${i === 0 ? \"border-left-color: var(--accent); font-weight:700; color:#fff;\" : \"color:#cbd5e1;\"}\">\n                        ${b}\n                    </div>\n                `).join(\"\");\n\n                const relatedBtns = relatedQs.map(rq => `\n                    <button onclick=\"jumpToQuestion(${rq.id})\" class=\"btn btn-ghost\" style=\"padding:3px 9px; font-size:0.76rem;\">\n                        Câu ${rq.id}\n                    </button>\n                `).join(\"\");\n\n                card.innerHTML = `\n                    <div>\n                        <div class=\"wf-card-header\">\n                            <span class=\"wf-title\">#${wf.id}. ${wf.root}</span>\n                            <button onclick=\"speakText(this.getAttribute('data-say'))\" data-say=\"${escHtml(wf.root.toLowerCase())}\" class=\"btn btn-ghost\" style=\"padding:4px 9px; font-size:0.78rem;\">🔊 Phát âm</button>\n                            <button onclick=\"openPhonetics('word', this.getAttribute('data-say'), '')\" data-say=\"${escHtml(wf.root.toLowerCase())}\" class=\"btn btn-ghost\" style=\"padding:4px 9px; font-size:0.78rem;\" title=\"Cách phát âm\">👄</button>\n                        </div>\n                        <div>${branchesHtml}</div>\n                    </div>\n                    <div style=\"margin-top:12px; padding-top:10px; border-top:1px solid var(--border); display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:6px;\">\n                        <span style=\"font-size:0.78rem; color:var(--text-muted);\">Bài tập trong đề:</span>\n                        <div style=\"display:flex; gap:5px; flex-wrap:wrap;\">${relatedBtns}</div>\n                    </div>\n                `;\n                container.appendChild(card);\n            });\n\n            // mục từ điển (dùng cho dạng Word entry)\n            const dc = el(\"de-grid-container\"), dt = el(\"de-title\");\n            dc.innerHTML = \"\";\n            const fe = entriesData.filter(e => (!scoped || questionsData.some(q => qType(q) === \"word_entry\" && q.entryId === e.id && inScope(q))) && (!query || (e.name + \" \" + (e.defs || []).join(\" \") + \" \" + (e.examples || []).join(\" \")).toLowerCase().includes(query)));\n            dt.classList.toggle(\"hidden\", !entriesData.length);\n            fe.forEach(e => {\n                const rel = questionsData.filter(q => qType(q) === \"word_entry\" && q.entryId === e.id);\n                const card = document.createElement(\"div\");\n                card.className = \"wf-card\";\n                card.innerHTML = entryHtml(e, true, false).replace('class=\"entry-card\"', 'style=\"max-width:none; margin:0; border:none; padding:0;\"') +\n                    `<div style=\"margin-top:12px; padding-top:10px; border-top:1px solid var(--border); display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:6px;\">\n                        <span style=\"font-size:0.78rem; color:var(--text-muted);\">Bài tập trong đề:</span>\n                        <div style=\"display:flex; gap:5px; flex-wrap:wrap;\">${rel.map(rq => `<button onclick=\"jumpToQuestion(${rq.id})\" class=\"btn btn-ghost\" style=\"padding:3px 9px; font-size:0.76rem;\">Câu ${rq.id}</button>`).join(\"\")}</div>\n                    </div>`;\n                dc.appendChild(card);\n            });\n        }\n\n        function jumpToQuestion(qId) {\n            clearScope(true);\n            resetFcFilterUI();\n            fillTypeSelect(\"fc-type-filter\"); fillTypeSelect(\"ws-type-filter\");\n            el(\"worksheet-list\").innerHTML = \"\";\n            appState.currentOrder = questionsData.map(q => q.id);\n            const idx = appState.currentOrder.indexOf(qId);\n            appState.currentIndex = idx >= 0 ? idx : 0;\n            switchMode(\"flashcard\");\n        }\n\n        // ==========================================================================\n        // 9. CHUYỂN CHẾ ĐỘ & RESET TIẾN ĐỘ\n        // ==========================================================================\n        function switchMode(mode) {\n            if (acGuardSwitch(mode)) return;\n            appState.currentMode = mode;\n            [\"flashcard\", \"worksheet\", \"handbook\"].forEach(m => {\n                document.getElementById(`section-${m}`).classList.toggle(\"hidden\", m !== mode);\n                document.getElementById(`tab-${m}`).classList.toggle(\"active\", m === mode);\n            });\n\n            if (mode === \"flashcard\") renderFlashcardView();\n            if (mode === \"worksheet\" && document.getElementById(\"worksheet-list\").children.length === 0) {\n                generateWorksheet(appState.wsQuizSize);\n            }\n            if (mode === \"handbook\") renderWordFamilies();\n            acOnMode(mode);\n        }\n\n        function openResetModal() {\n            document.getElementById(\"reset-modal\").classList.remove(\"hidden\");\n        }\n\n        function closeResetModal() {\n            document.getElementById(\"reset-modal\").classList.add(\"hidden\");\n        }\n\n        function confirmResetAll() {\n            syncReset();\n            appState.answers = {};\n            appState.streak = 0;\n            appState.currentIndex = 0;\n            resetFcFilterUI();\n            appState.currentOrder = scopedIds();\n            saveState();\n            updateStatsUI();\n            renderFlashcardView();\n            generateWorksheet(appState.wsQuizSize);\n            closeResetModal();\n        }\n\n        // Khởi chạy ứng dụng khi tải trang\n// ==========================================================================\n        // 10. BỔ SUNG: ĐĂNG NHẬP & TỰ GỬI KẾT QUẢ CHO GIÁO VIÊN\n        //     (phần gốc ở trên giữ nguyên; phần này chỉ nối ứng dụng với máy chủ)\n        // ==========================================================================\n        // >>> Dán URL Web App của Google Apps Script (kết thúc bằng /exec) vào dòng dưới đây <<<\n        const API_URL = \"DÁN_URL_WEB_APP_VÀO_ĐÂY\";\n        const PRELOAD_CLASSES = null;   // máy chủ tự điền danh sách lớp khi phát trang\n        const SESSION_KEY = \"wf9_nmc_session_v4\";\n        const S2L = { \"Đ\": \"correct\", \"S\": \"wrong\", \"X\": \"revealed\", \"T+\": \"correct\", \"T-\": \"wrong\" };\n        const L2S = { correct: \"Đ\", wrong: \"S\", revealed: \"X\" };\n\n        let session = null;          // { cls, name, token }\n        let progressSeed = null;     // kết quả đã lưu trên máy chủ\n        let bankSeed = null;         // câu hỏi do máy chủ trả kèm khi đăng nhập\n        let syncQueue = {};          // qid -> { qid, text, revealed, mode }   (đang chờ gửi)\n        let lastSent = {};           // qid -> chữ ký đã gửi/nhận từ máy chủ\n        let pendingReset = false;    // đã bấm \"Làm lại từ đầu\" mà máy chủ chưa xóa\n        let syncing = false, syncTimer = null, syncFails = 0;\n\n        function el(id) { return document.getElementById(id); }\n        function escHtml(str) {\n            return String(str == null ? \"\" : str).replace(/&/g, \"&amp;\").replace(/</g, \"&lt;\").replace(/>/g, \"&gt;\").replace(/\"/g, \"&quot;\").replace(/'/g, \"&#39;\");\n        }\n        function newBatchId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }\n        function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }\n        function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }\n        function lsDel(k) { try { localStorage.removeItem(k); } catch (e) {} }\n\n        function showBox(id, msg) {\n            const box = el(id);\n            if (!box) return;\n            box.innerText = msg || \"\";\n            box.classList.toggle(\"hidden\", !msg);\n        }\n        function showNotice(message, onOk, withCancel) {\n            const modal = el(\"notice-modal\");\n            el(\"notice-text\").innerText = message;\n            el(\"notice-cancel\").classList.toggle(\"hidden\", !withCancel);\n            modal.classList.remove(\"hidden\");\n            el(\"notice-ok\").onclick = () => { modal.classList.add(\"hidden\"); if (onOk) onOk(); };\n            el(\"notice-cancel\").onclick = () => { modal.classList.add(\"hidden\"); };\n        }\n        function setLoading(on, text) {\n            el(\"loading-overlay\").classList.toggle(\"hidden\", !on);\n            if (text) el(\"loading-text\").innerText = text;\n        }\n\n        // ---------- gọi máy chủ ----------\n        function apiConfigured() { return /^https:\\/\\/script\\.google\\.com\\/.+\\/exec$/.test(API_URL); }\n\n        async function apiFetch(url, options) {\n            const ctrl = new AbortController();\n            const timer = setTimeout(() => ctrl.abort(), 25000);\n            try {\n                const res = await fetch(url, Object.assign({ redirect: \"follow\", signal: ctrl.signal }, options || {}));\n                return await res.json();\n            } finally { clearTimeout(timer); }\n        }\n        function apiGet(params) { return apiFetch(API_URL + \"?\" + new URLSearchParams(params).toString()); }\n        function apiPost(payload) {\n            return apiFetch(API_URL, { method: \"POST\", headers: { \"Content-Type\": \"text/plain;charset=utf-8\" }, body: JSON.stringify(payload) });\n        }\n        async function api(action, extra) {\n            const r = await apiPost(Object.assign({ action: action, cls: session.cls, name: session.name, token: session.token }, extra || {}));\n            if (r && r.error === \"AUTH\") {\n                forceLogout(\"Phiên đăng nhập không còn hợp lệ (có thể giáo viên đã chỉnh danh sách lớp). Hãy đăng nhập lại.\");\n                throw new Error(\"AUTH\");\n            }\n            return r;\n        }\n\n        // ---------- hàng chờ gửi kết quả ----------\n        function queueKey() { return STORAGE_KEY + \":queue\"; }\n        function persistQueue() { lsSet(queueKey(), JSON.stringify({ q: syncQueue, reset: pendingReset })); }\n        function loadQueue() {\n            syncQueue = {}; pendingReset = false;\n            try {\n                const o = JSON.parse(lsGet(queueKey()) || \"null\");\n                if (o && typeof o === \"object\") { syncQueue = o.q || {}; pendingReset = !!o.reset; }\n            } catch (e) {}\n        }\n\n        function sendText(rec) { return rec.blank ? \"\" : (rec.userText || \"\"); }\n        function sigOf(rec) { return rec.status + \"|\" + sendText(rec); }\n\n        function setSyncUI() {\n            const n = Object.keys(syncQueue).length + (pendingReset ? 1 : 0);\n            const e = el(\"sync-state\");\n            if (!e) return;\n            if (syncing) { e.className = \"sync busy\"; e.innerText = \"⏳ đang gửi\"; }\n            else if (n && syncFails) { e.className = \"sync err\"; e.innerText = \"⚠ chưa gửi được, sẽ tự gửi lại\"; }\n            else if (n) { e.className = \"sync busy\"; e.innerText = \"⏳ chờ gửi\"; }\n            else { e.className = \"sync\"; e.innerText = \"✔ đã gửi cho giáo viên\"; }\n        }\n\n        // so sánh câu trả lời hiện tại với lần gửi trước; câu nào mới/đổi thì xếp hàng chờ gửi\n        function queueSync() {\n            if (!session) return;\n            const mode = appState.currentMode === \"worksheet\" ? \"worksheet\" : \"flashcard\";\n            Object.keys(appState.answers).forEach(k => {\n                const rec = appState.answers[k];\n                if (!rec || !L2S[rec.status]) return;\n                if (lastSent[k] === sigOf(rec)) { delete syncQueue[k]; return; }\n                syncQueue[k] = { qid: parseInt(k, 10), text: sendText(rec), revealed: rec.status === \"revealed\", mode: mode, sig: sigOf(rec) };\n            });\n            persistQueue();\n            setSyncUI();\n            scheduleSync(500);\n        }\n\n        function scheduleSync(ms) {\n            clearTimeout(syncTimer);\n            syncTimer = setTimeout(flushQueue, ms);\n        }\n\n        async function flushQueue() {\n            if (!session || syncing) return;\n            if (!pendingReset && !Object.keys(syncQueue).length) { setSyncUI(); return; }\n            syncing = true; setSyncUI();\n            try {\n                if (pendingReset) {\n                    const r = await api(\"resetme\");\n                    if (!r || !r.ok) throw new Error(\"SERVER\");\n                    pendingReset = false; persistQueue();\n                }\n                for (const mode of [\"flashcard\", \"worksheet\"]) {\n                    const items = Object.keys(syncQueue).map(k => syncQueue[k]).filter(i => i.mode === mode);\n                    if (!items.length) continue;\n                    const r = await api(\"submit\", {\n                        mode: mode, batchId: newBatchId(),\n                        items: items.map(i => ({ qid: i.qid, text: i.text, revealed: i.revealed }))\n                    });\n                    if (!r || !r.ok) throw new Error((r && r.error) || \"SERVER\");\n                    items.forEach(i => {\n                        lastSent[i.qid] = i.sig;\n                        const cur = syncQueue[i.qid];\n                        if (cur && cur.sig === i.sig) delete syncQueue[i.qid];\n                    });\n                    persistQueue();\n                }\n                syncFails = 0;\n            } catch (e) {\n                if (e && e.message === \"AUTH\") { syncing = false; return; }\n                syncFails++;\n            }\n            syncing = false;\n            setSyncUI();\n            if (pendingReset || Object.keys(syncQueue).length) scheduleSync(Math.min(30000, 2000 * Math.pow(2, Math.min(syncFails, 4))));\n        }\n\n        function syncReset() {\n            syncQueue = {}; lastSent = {}; pendingReset = true;\n            persistQueue(); setSyncUI();\n            scheduleSync(200);\n        }\n\n        window.addEventListener(\"online\", () => { if (session) scheduleSync(300); });\n        document.addEventListener(\"visibilitychange\", () => { if (document.hidden && session) flushQueue(); });\n\n        // ---------- đăng nhập ----------\n        function loginErrorText(code) {\n            switch (code) {\n                case \"NOT_FOUND\": return \"Không tìm thấy bạn trong danh sách của lớp này. Hãy kiểm tra lại LỚP và HỌ TÊN (gõ có dấu, đúng như danh sách lớp). Nếu vẫn không vào được, hãy báo giáo viên.\";\n                case \"AMBIGUOUS\": return \"Trong lớp có hai bạn tên gần giống nhau. Hãy gõ ĐẦY ĐỦ DẤU tiếng Việt để hệ thống phân biệt.\";\n                case \"DUPLICATE_ROSTER\": return \"Danh sách lớp đang có hai bạn trùng họ tên. Hãy báo giáo viên để chỉnh lại danh sách.\";\n                case \"NEED_CODE\": return \"Bạn này có mã học sinh. Hãy nhập mã học sinh do giáo viên cấp.\";\n                case \"BAD_CODE\": return \"Mã học sinh chưa đúng.\";\n                default: return \"Không kết nối được máy chủ. Kiểm tra mạng rồi thử lại.\";\n            }\n        }\n\n        async function loadClassList() {\n            const sel = el(\"login-class\");\n            try {\n                const r = Array.isArray(PRELOAD_CLASSES) ? { ok: true, classes: PRELOAD_CLASSES } : await apiGet({ action: \"classes\" });\n                if (!r.ok) throw new Error(\"server\");\n                if (!r.classes || !r.classes.length) {\n                    sel.innerHTML = '<option value=\"\">(Chưa có lớp nào)</option>';\n                    showBox(\"login-error\", \"Giáo viên chưa nhập danh sách lớp. Hãy báo giáo viên.\");\n                    return;\n                }\n                sel.innerHTML = '<option value=\"\">-- Chọn lớp --</option>' + r.classes.map(c => `<option value=\"${escHtml(c)}\">${escHtml(c)}</option>`).join(\"\");\n                if (r.classes.length === 1) {   // link riêng của một lớp: chọn sẵn lớp, đưa con trỏ vào ô họ tên\n                    sel.value = r.classes[0];\n                    const nm = el(\"login-name\"); if (nm) { try { nm.focus(); } catch (e2) {} }\n                }\n            } catch (e) {\n                sel.innerHTML = '<option value=\"\">(Chưa tải được danh sách lớp)</option>';\n                showBox(\"login-error\", \"Không tải được danh sách lớp. Kiểm tra mạng rồi tải lại trang.\");\n            }\n        }\n\n        async function submitLogin(ev) {\n            if (ev && ev.preventDefault) ev.preventDefault();\n            const cls = el(\"login-class\").value;\n            const name = el(\"login-name\").value.trim();\n            const code = el(\"login-code\").value.trim();\n            const btn = el(\"login-btn\");\n            if (!cls) { showBox(\"login-error\", \"Hãy chọn lớp của bạn.\"); return false; }\n            if (!name) { showBox(\"login-error\", \"Hãy nhập họ và tên của bạn.\"); return false; }\n            showBox(\"login-error\", \"\");\n            btn.disabled = true; btn.innerText = \"Đang kiểm tra...\";\n            try {\n                const r = await apiPost({ action: \"login\", cls: cls, name: name, code: code });\n                if (r && r.ok) {\n                    session = { cls: r.cls, name: r.name, token: r.token };\n                    lsSet(SESSION_KEY, JSON.stringify(session));\n                    progressSeed = r.progress || null;\n                    bankSeed = r.bank || null; setAcCfg(r.ac);\n                    await startApp();\n                } else {\n                    const err = r && r.error;\n                    if (err === \"NEED_CODE\" || err === \"BAD_CODE\") el(\"login-code-wrap\").classList.remove(\"hidden\");\n                    showBox(\"login-error\", loginErrorText(err));\n                }\n            } catch (e) {\n                showBox(\"login-error\", loginErrorText(\"NETWORK\"));\n            } finally {\n                btn.disabled = false; btn.innerText = \"Vào làm bài ➜\";\n            }\n            return false;\n        }\n\n        function clearSessionAndReload() { lsDel(SESSION_KEY); location.reload(); }\n        function logoutStudent() {\n            const n = Object.keys(syncQueue).length;\n            const extra = n ? \" (Còn \" + n + \" câu chưa gửi được; sẽ tự gửi khi bạn đăng nhập lại trên máy này.)\" : \"\";\n            showNotice(\"Đăng xuất khỏi tài khoản \" + session.name + \"?\" + extra, clearSessionAndReload, true);\n        }\n        function forceLogout(message) { showNotice(message, clearSessionAndReload, false); }\n\n        // ---------- khởi chạy sau đăng nhập ----------\n        function refreshCounts() {\n            const n = questionsData.length;\n            document.querySelectorAll(\".n-q\").forEach(e => { e.innerText = n; });\n            el(\"fc-filter-all\").innerText = \"Tất cả (\" + n + \")\";\n            el(\"ws-size-all\").innerText = \"📋 Tất cả \" + Math.min(n, WS_ALL_MAX) + \" câu\";\n            fillTypeSelect(\"fc-type-filter\");\n            fillTypeSelect(\"ws-type-filter\");\n            const pw = el(\"progress-wrap\");\n            if (pw) pw.title = \"Tiến độ hoàn thành \" + n + \" câu\";\n            el(\"ws-score-detail\").innerText = \"Đúng 0 / \" + n + \" câu (0%)\";\n        }\n\n        // kết quả máy chủ -> dạng answers của file gốc\n        function answersFromServer() {\n            const out = {};\n            const p = progressSeed || {};\n            Object.keys(p.states || {}).forEach(k => {\n                const st = S2L[p.states[k]];\n                if (!st || !getQuestionById(parseInt(k, 10))) return;\n                out[k] = { userText: (p.texts && p.texts[k]) || \"\", status: st };\n            });\n            return out;\n        }\n\n        async function startApp() {\n            setLoading(true, \"Đang tải câu hỏi...\");\n            try {\n                // đăng nhập / whoami đã trả sẵn câu hỏi; chỉ gọi riêng khi thiếu\n                let r = bankSeed;\n                bankSeed = null;\n                if (!r) r = await api(\"bank\");\n                if (!r || (r.ok === false)) throw new Error((r && r.error) || \"SERVER\");\n                questionsData = r.questions || [];\n                wordFamilies = r.families || [];\n                entriesData = r.entries || [];\n            } catch (e) {\n                setLoading(false);\n                if (e.message !== \"AUTH\") showBox(\"login-error\", \"Đăng nhập được nhưng chưa tải được câu hỏi. Kiểm tra mạng rồi thử lại.\");\n                return;\n            }\n            setLoading(false);\n\n            STORAGE_KEY = \"wf9_nmc_progress_v1:\" + session.cls + \":\" + session.name;\n            acLoad();\n            acOnStart();\n            loadQueue();\n            // máy chủ là nguồn đúng; riêng câu còn chờ gửi thì giữ bản trên máy này\n            let local = {};\n            try { local = (JSON.parse(lsGet(STORAGE_KEY) || \"null\") || {}).answers || {}; } catch (e) {}\n            const server = pendingReset ? {} : answersFromServer();\n            loadState();                         // chỉ lấy streak + vị trí câu đang làm\n            appState.answers = server;\n            Object.keys(syncQueue).forEach(k => { if (local[k]) appState.answers[k] = local[k]; });\n            lastSent = {};\n            Object.keys(server).forEach(k => { lastSent[k] = sigOf(server[k]); });\n\n            appState.fcFilter = \"all\";\n            appState.scope = { unit: \"\", topic: \"\", q: \"\" };\n            if (el(\"scope-search\")) el(\"scope-search\").value = \"\";\n            appState.currentOrder = questionsData.map(q => q.id);\n            if (!(appState.currentIndex >= 0 && appState.currentIndex < appState.currentOrder.length)) appState.currentIndex = 0;\n            appState.wsQuizSize = 1000;\n            appState.wsQuestionIds = questionsData.map(q => q.id);\n            appState.wsSubmitted = false;\n            [\"all\", \"unanswered\", \"wrong\"].forEach(f => { el(\"fc-filter-\" + f).className = (f === \"all\") ? \"btn btn-accent\" : \"btn btn-ghost\"; });\n            [\"fc-type-filter\", \"ws-type-filter\"].forEach(id => { const e = el(id); if (e) e.value = \"all\"; });\n            el(\"worksheet-list\").innerHTML = \"\";\n\n            el(\"login-screen\").classList.add(\"hidden\");\n            el(\"app-root\").classList.remove(\"hidden\");\n            el(\"student-chip-text\").innerText = session.name + \" · \" + session.cls;\n            showBox(\"app-error\", questionsData.length ? \"\" : \"Giáo viên chưa nhập câu hỏi nào. Hãy báo giáo viên.\");\n            refreshCounts();\n            renderScopeBar();\n            setSyncUI();\n\n            updateStatsUI();\n            renderFlashcardView();\n            appState.wsQuizSize = 1000;\n            if (appState.currentMode === \"worksheet\") generateWorksheet(1000);\n            else acOnMode(appState.currentMode);   // như trước: cập nhật thanh theo dõi ngay khi vào\n            if (appState.currentMode === \"handbook\") renderWordFamilies();\n            if (pendingReset || Object.keys(syncQueue).length) scheduleSync(300);\n        }\n\n        async function bootstrapAuth() {\n            if (!apiConfigured()) {\n                showBox(\"login-error\", \"Trang này chưa được kết nối với giáo viên (thiếu API_URL). Hãy báo giáo viên.\");\n                el(\"login-btn\").disabled = true;\n                el(\"login-class\").innerHTML = '<option value=\"\">(Chưa cấu hình)</option>';\n                return;\n            }\n            let saved = null;\n            try { saved = JSON.parse(lsGet(SESSION_KEY) || \"null\"); } catch (e) {}\n            if (saved && saved.token) {\n                try {\n                    const r = await apiPost({ action: \"whoami\", cls: saved.cls, name: saved.name, token: saved.token });\n                    if (r && r.ok) {\n                        session = { cls: r.cls, name: r.name, token: saved.token };\n                        progressSeed = r.progress || null;\n                        bankSeed = r.bank || null; setAcCfg(r.ac);\n                        await startApp();\n                        if (!el(\"app-root\").classList.contains(\"hidden\")) return;\n                    } else { lsDel(SESSION_KEY); }\n                } catch (e) {\n                    showBox(\"login-error\", \"Không kết nối được máy chủ để xác thực. Kiểm tra mạng rồi đăng nhập lại.\");\n                }\n            }\n            loadClassList();\n        }\n\n                document.addEventListener(\"keydown\", ev => {\n            if (!session || el(\"app-root\").classList.contains(\"hidden\") || el(\"section-flashcard\").classList.contains(\"hidden\")) return;\n            const tag = (ev.target && ev.target.tagName) || \"\";\n            if (tag === \"INPUT\" || tag === \"TEXTAREA\" || tag === \"SELECT\" || ev.ctrlKey || ev.metaKey || ev.altKey) return;\n            if (!el(\"reset-modal\").classList.contains(\"hidden\") || !el(\"ph-modal\").classList.contains(\"hidden\")) return;\n            const q = getCurrentQuestion();\n            if (!q || qType(q) !== \"trac_nghiem\") return;\n            const k = ev.key.toUpperCase();\n            const L = /^[A-D]$/.test(k) ? k : (/^[1-4]$/.test(k) ? OPT_LETTERS[+k - 1] : \"\");\n            if (L) { selectFcOption(L); ev.preventDefault(); }\n            else if (ev.key === \"Enter\" && (!tag || tag === \"BODY\" || tag === \"HTML\" || (ev.target.classList && ev.target.classList.contains(\"opt-btn\")))) { handleFlashcardEnter(); ev.preventDefault(); }\n        });\n\n\n        // ==========================================================================\n        // 11. CHỐNG GIAN LẬN (v8): phát hiện rời trang, nhắc nhở, tự nộp, ghi cho giáo viên\n        //     Giáo viên chọn cho từng lớp: free (tự do) | exam (giám sát khi Làm đề tổng hợp)\n        //     | all (giám sát cả khi luyện tập). Máy chủ gửi cài đặt kèm lúc đăng nhập.\n        // ==========================================================================\n        let acCfg = { mode: \"free\", limit: 4 };\n        let acData = null;        // theo từng học sinh, lưu trên máy: { ws: {id,n,secs,open,done}, pr: {n,secs}, q: [] }\n        let acEp = null;          // lần rời trang đang diễn ra { start, where }\n        let acQ = [];             // sự kiện chờ gửi\n        let acFlushing = false, acRetry = null, acBypass = false;\n\n        function setAcCfg(c) {\n            const m = c && (c.mode === \"exam\" || c.mode === \"all\") ? c.mode : \"free\";\n            let lim = c ? parseInt(c.limit, 10) : 4;\n            if (!(lim >= 1)) lim = 4;\n            acCfg = { mode: m, limit: Math.min(20, lim) };\n        }\n        function acKey() { return \"wf9_nmc_ac_v1:\" + session.cls + \":\" + session.name; }\n        function acLoad() {\n            let o = null;\n            try { o = JSON.parse(lsGet(acKey()) || \"null\"); } catch (e) {}\n            if (!o || typeof o !== \"object\") o = {};\n            acData = { ws: o.ws && typeof o.ws === \"object\" ? o.ws : null, pr: o.pr && typeof o.pr === \"object\" ? o.pr : { n: 0, secs: 0 }, q: Array.isArray(o.q) ? o.q : [] };\n            acData.pr.n = acData.pr.n || 0; acData.pr.secs = acData.pr.secs || 0;\n            acQ = acData.q;\n        }\n        function acSave() { if (!acData || !session) return; acData.q = acQ.slice(-50); acQ = acData.q; lsSet(acKey(), JSON.stringify(acData)); }\n        function acOn() { return !!session && !!acData && acCfg.mode !== \"free\"; }\n        function acEnsureWs() {\n            if (!acData.ws || acData.ws.done) acData.ws = { id: newBatchId(), n: 0, secs: 0, open: false, done: false };\n            return acData.ws;\n        }\n        // đang ở chế độ nào cần theo dõi: \"worksheet\" | \"practice\" | null\n        function acWhere() {\n            if (!acOn()) return null;\n            if (appState.currentMode === \"worksheet\" && !appState.wsSubmitted) return \"worksheet\";\n            return acCfg.mode === \"all\" ? \"practice\" : null;\n        }\n\n        function acNotice(msg) {\n            const t = el(\"notice-text\"); if (t) t.style.whiteSpace = \"pre-line\";\n            showNotice(msg, null, false);\n        }\n\n        function acUpdateBar() {\n            const b = el(\"ac-bar\"); if (!b) return;\n            let msg = \"\", bad = false;\n            const w = acWhere();\n            if (w === \"worksheet\") {\n                const n = acData && acData.ws ? acData.ws.n : 0;\n                msg = \"🛡 Đề đang được giám sát: bạn đã rời trang \" + n + \"/\" + acCfg.limit + \" lần. Rời trang quá \" + acCfg.limit + \" lần thì bài sẽ tự nộp. Đừng chuyển tab hoặc mở ứng dụng khác khi đang làm bài.\";\n                bad = n >= acCfg.limit;\n            } else if (w === \"practice\") {\n                msg = \"🛡 Giáo viên đang theo dõi việc rời trang khi luyện tập: đã ghi nhận \" + acData.pr.n + \" lần (chỉ ghi nhận, không bị tự nộp).\";\n            }\n            b.textContent = msg;\n            b.classList.toggle(\"hidden\", !msg);\n            b.classList.toggle(\"bad\", bad);\n        }\n\n        function acOnMode(mode) {\n            if (!acData) return;\n            const ws = el(\"section-worksheet\");\n            if (ws) ws.classList.toggle(\"ac-lock\", acOn());\n            if (acOn() && mode === \"worksheet\" && !appState.wsSubmitted) { acEnsureWs().open = true; acSave(); }\n            else if (acData.ws && acData.ws.open) { acData.ws.open = false; acSave(); }\n            acUpdateBar();\n        }\n\n        // học sinh bấm sang chế độ khác khi đang làm đề có giám sát: tính là một lần rời trang\n        function acGuardSwitch(mode) {\n            if (acBypass || !acOn() || mode === \"worksheet\" || appState.currentMode !== \"worksheet\" || appState.wsSubmitted) return false;\n            if (!acData.ws || acData.ws.done) return false;\n            const n = acData.ws.n;\n            const t = el(\"notice-text\"); if (t) t.style.whiteSpace = \"pre-line\";\n            showNotice(\"Bạn đang làm đề có giám sát.\\nChuyển sang chế độ khác sẽ tính là 1 lần rời trang (hiện \" + n + \"/\" + acCfg.limit + \"), và bạn có thể thấy đáp án khi luyện tập.\\nBạn có muốn tiếp tục?\", () => {\n                const auto = acRecordLeave(\"worksheet\", 0, \"Chuyển sang chế độ khác\");\n                if (!auto) { acBypass = true; try { switchMode(mode); } finally { acBypass = false; } }\n            }, true);\n            return true;\n        }\n\n        function acOnStart() {\n            if (!acOn()) return;\n            const a = acData.ws;\n            if (a && !a.done && a.open) {      // đang làm đề mà tải lại / đóng rồi mở lại trang\n                a.open = false;\n                acRecordLeave(\"worksheet\", 0, \"Tải lại hoặc đóng trang khi đang làm đề\");\n            }\n            acFlush();\n        }\n\n        function acOnSubmit(opts) {\n            appState.wsSubmitted = true;\n            if (acData && acData.ws && !acData.ws.done) { acData.ws.done = true; acData.ws.open = false; acSave(); }\n            if (acOn()) document.querySelectorAll(\"#worksheet-list input[type=text], #worksheet-list textarea\").forEach(e => { e.readOnly = true; });\n            if (opts && opts.auto) {\n                const d = el(\"ws-score-detail\");\n                if (d) d.innerText += \" · Bài được tự nộp vì rời trang quá \" + acCfg.limit + \" lần\";\n            }\n            acUpdateBar();\n        }\n\n        // ghi nhận một lần rời trang. Trả true nếu bài bị tự nộp / kết thúc.\n        function acRecordLeave(where, secs, note) {\n            let n, auto = false;\n            if (where === \"worksheet\") {\n                const a = acEnsureWs();\n                a.n++; a.secs += secs; n = a.n;\n                auto = a.n > acCfg.limit;\n            } else {\n                acData.pr.n++; acData.pr.secs += secs; n = acData.pr.n;\n            }\n            acQ.push({ ev: newBatchId(), where: where, n: n, secs: secs, auto: auto, note: note || \"\" });\n            acSave();\n            acFlush();\n            if (where === \"worksheet\") {\n                if (auto) {\n                    if (appState.currentMode === \"worksheet\" && !appState.wsSubmitted) {\n                        submitWorksheetQuiz({ auto: true });\n                        acNotice(\"⛔ Bạn đã rời trang \" + n + \" lần, quá \" + acCfg.limit + \" lần được phép.\\nBài làm đã được tự động nộp và chấm điểm. Giáo viên sẽ thấy số lần rời trang của bạn.\");\n                    } else {\n                        if (acData.ws) { acData.ws.done = true; acData.ws.open = false; acSave(); }\n                        acNotice(\"⛔ Bài làm trước đó đã kết thúc vì bạn rời trang \" + n + \" lần, quá \" + acCfg.limit + \" lần được phép. Giáo viên sẽ thấy việc này.\");\n                    }\n                } else {\n                    const left = acCfg.limit - n;\n                    acNotice(\"⚠️ Bạn vừa rời khỏi trang làm bài\" + (secs ? \" (\" + secs + \" giây)\" : \"\") + \".\\nĐây là lần thứ \" + n + \"/\" + acCfg.limit + \" được phép. \" +\n                        (left > 0 ? \"Còn \" + left + \" lần; rời trang thêm quá \" + acCfg.limit + \" lần thì bài tự nộp.\" : \"Đây là lần cuối: lần rời trang tiếp theo bài sẽ tự nộp.\") +\n                        \"\\nGiáo viên xem được số lần rời trang. Hãy ở yên trong trang khi làm bài.\");\n                }\n            }\n            acUpdateBar();\n            return auto;\n        }\n\n        async function acFlush() {\n            if (!session || !acQ.length || acFlushing) return;\n            acFlushing = true;\n            try {\n                while (acQ.length) {\n                    const e = acQ[0];\n                    const r = await api(\"ac_event\", { ev: e.ev, where: e.where, n: e.n, secs: e.secs, auto: e.auto, note: e.note });\n                    if (!r || !r.ok) throw new Error(\"SERVER\");\n                    acQ.shift(); acSave();\n                }\n            } catch (e) {\n                clearTimeout(acRetry);\n                acRetry = setTimeout(acFlush, 10000);\n            }\n            acFlushing = false;\n        }\n\n        function acBegin() {\n            if (acEp) return;\n            const w = acWhere();\n            if (!w) return;\n            acEp = { start: Date.now(), where: w };\n        }\n        function acEnd() {\n            if (!acEp) return;\n            const ep = acEp; acEp = null;\n            const ms = Date.now() - ep.start;\n            if (ms < 1500) return;                 // bỏ qua chớp mắt (thông báo kéo xuống, chạm nhầm)\n            acRecordLeave(ep.where, Math.round(ms / 1000), \"\");\n        }\n        document.addEventListener(\"visibilitychange\", () => { if (document.hidden) acBegin(); else acEnd(); });\n        window.addEventListener(\"blur\", () => { setTimeout(() => { if (!document.hidden && !document.hasFocus()) acBegin(); }, 700); });\n        window.addEventListener(\"focus\", () => { if (!document.hidden) acEnd(); });\n        window.addEventListener(\"online\", () => { if (session) acFlush(); });\n\n        // khi làm đề có giám sát: không cho sao chép / dán / chuột phải trong vùng làm đề\n        function acInWs(ev) {\n            let t = ev.target; if (t && t.nodeType !== 1) t = t.parentElement;\n            return !!(t && t.closest && t.closest(\"#section-worksheet\"));\n        }\n        [\"copy\", \"cut\", \"contextmenu\", \"dragstart\", \"paste\"].forEach(type => {\n            document.addEventListener(type, ev => { if (acOn() && acWhere() === \"worksheet\" && acInWs(ev)) ev.preventDefault(); }, true);\n        });\n\n        window.onload = bootstrapAuth;\n    </script>\n</body>\n</html>";
var PAGE_DASHBOARD_ = "<!DOCTYPE html>\n<html lang=\"vi\">\n<head>\n    <meta charset=\"UTF-8\">\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n    <title>Bảng điều khiển giáo viên - Word Form 9 by NMC</title>\n    <style>\n        :root {\n            --bg-dark: #0f172a;\n            --bg-card: #1e293b;\n            --text-main: #f8fafc;\n            --text-muted: #94a3b8;\n            --primary: #38bdf8;\n            --accent: #fbbf24;\n            --success: #22c55e;\n            --danger: #f43f5e;\n            --border: rgba(255, 255, 255, 0.12);\n            --radius-lg: 18px;\n            --radius-md: 12px;\n            --radius-sm: 8px;\n            --shadow: 0 10px 30px -5px rgba(0, 0, 0, 0.45);\n            --font-sans: system-ui, -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;\n        }\n\n        * { box-sizing: border-box; margin: 0; padding: 0; }\n\n        body {\n            font-family: var(--font-sans);\n            background: radial-gradient(circle at top, #172554 0%, #0f172a 55%, #090d16 100%);\n            background-attachment: fixed;\n            color: var(--text-main);\n            min-height: 100vh;\n            line-height: 1.55;\n            padding-bottom: 60px;\n        }\n\n        .container { max-width: 1280px; margin: 0 auto; padding: 16px; }\n        .hidden { display: none !important; }\n\n        .panel {\n            background: rgba(30, 41, 59, 0.88);\n            border: 1px solid var(--border);\n            border-radius: var(--radius-lg);\n            padding: 20px;\n            box-shadow: var(--shadow);\n            margin-bottom: 16px;\n        }\n\n        .panel h2 { font-size: 1.05rem; font-weight: 800; margin-bottom: 12px; }\n\n        .app-header {\n            display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 12px;\n            background: rgba(30, 41, 59, 0.88);\n            border: 1px solid var(--border);\n            border-radius: var(--radius-lg);\n            padding: 14px 20px;\n            margin-bottom: 16px;\n            box-shadow: var(--shadow);\n        }\n\n        .brand { display: flex; align-items: center; gap: 12px; }\n        .logo {\n            width: 44px; height: 44px; border-radius: 13px;\n            background: linear-gradient(135deg, #fbbf24, #f59e0b);\n            display: flex; align-items: center; justify-content: center; font-size: 22px; flex-shrink: 0;\n        }\n        .brand h1 { font-size: 1.3rem; font-weight: 800; letter-spacing: -0.02em; }\n        .brand small { color: var(--text-muted); font-size: 0.8rem; }\n\n        .btn {\n            border: none; border-radius: var(--radius-sm); padding: 9px 14px; font-size: 0.88rem; font-weight: 700;\n            cursor: pointer; font-family: inherit; display: inline-flex; align-items: center; gap: 6px;\n            transition: background 0.15s;\n        }\n        .btn:disabled { opacity: 0.6; cursor: default; }\n        .btn-accent { background: var(--accent); color: #0f172a; }\n        .btn-accent:hover:not(:disabled) { background: #fcd34d; }\n        .btn-ghost { background: rgba(255, 255, 255, 0.09); color: var(--text-main); border: 1px solid var(--border); }\n        .btn-ghost:hover { background: rgba(255, 255, 255, 0.16); }\n\n        .field, select, input[type=text], input[type=password] {\n            background: rgba(15, 23, 42, 0.9);\n            border: 1px solid var(--border);\n            color: #fff;\n            border-radius: var(--radius-sm);\n            padding: 9px 12px;\n            font-size: 0.92rem;\n            font-family: inherit;\n            outline: none;\n        }\n        select:focus, input:focus { border-color: var(--primary); }\n\n        /* Cổng mật khẩu */\n        .gate { min-height: 80vh; display: flex; align-items: center; justify-content: center; }\n        .gate-card { width: 100%; max-width: 400px; }\n        .gate-card input { width: 100%; margin-bottom: 12px; padding: 12px 14px; font-size: 1rem; }\n        .error-box {\n            background: rgba(244, 63, 94, 0.14); border: 1px solid rgba(244, 63, 94, 0.5);\n            color: #fecdd3; border-radius: 10px; padding: 10px 12px; font-size: 0.9rem; margin-bottom: 12px;\n        }\n\n        /* Thanh lọc */\n        .toolbar { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }\n        .toolbar .grow { flex: 1; min-width: 200px; }\n        .muted { color: var(--text-muted); font-size: 0.84rem; }\n\n        /* KPI */\n        .kpi-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 10px; margin-bottom: 16px; }\n        .kpi {\n            background: rgba(30, 41, 59, 0.88); border: 1px solid var(--border);\n            border-radius: var(--radius-md); padding: 14px 12px; text-align: center;\n        }\n        .kpi .label { font-size: 0.72rem; color: var(--text-muted); font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; }\n        .kpi .val { font-size: 1.7rem; font-weight: 800; margin-top: 2px; }\n        .kpi .sub { font-size: 0.75rem; color: var(--text-muted); }\n\n        /* Bảng */\n        .table-wrap { overflow-x: auto; border-radius: var(--radius-md); border: 1px solid var(--border); }\n        table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }\n        th, td { padding: 9px 12px; text-align: left; white-space: nowrap; border-bottom: 1px solid rgba(255, 255, 255, 0.07); }\n        th {\n            background: rgba(15, 23, 42, 0.95); font-size: 0.74rem; text-transform: uppercase; letter-spacing: 0.04em;\n            color: var(--text-muted); cursor: pointer; user-select: none; position: sticky; top: 0;\n        }\n        th:hover { color: #fff; }\n        th.sorted { color: var(--accent); }\n        tbody tr:hover { background: rgba(255, 255, 255, 0.04); }\n        tr.clickable { cursor: pointer; }\n        td.num { text-align: right; font-variant-numeric: tabular-nums; }\n\n        .bar { width: 110px; height: 8px; background: rgba(255, 255, 255, 0.1); border-radius: 99px; overflow: hidden; display: inline-block; vertical-align: middle; }\n        .bar > i { display: block; height: 100%; background: linear-gradient(90deg, #38bdf8, #22c55e); }\n        .bar.bad > i { background: linear-gradient(90deg, #fbbf24, #f43f5e); }\n\n        .strip { display: inline-flex; gap: 1px; vertical-align: middle; }\n        .strip i { width: 6px; height: 14px; border-radius: 1px; background: rgba(255, 255, 255, 0.1); }\n        .strip i.c { background: #22c55e; }\n        .strip i.w { background: #f43f5e; }\n        .strip i.r { background: #fbbf24; }\n        .strip i.p { background: #cbd5e1; }\n        .scope-chip { display: inline-block; font-size: 0.72rem; padding: 1px 8px; border-radius: 999px; background: rgba(56,189,248,0.16); color: #7dd3fc; margin: 1px 3px 1px 0; white-space: nowrap; }\n        .chk-grid { display: flex; flex-wrap: wrap; gap: 6px 14px; margin: 6px 0 10px; }\n        .chk-grid label { display: flex; gap: 6px; align-items: center; cursor: pointer; font-size: 0.9rem; }\n        .link-url { font-family: monospace; font-size: 0.78rem; word-break: break-all; color: var(--text-muted, #94a3b8); }\n        .strip i.sg { background: #22c55e; opacity: .55; border-bottom: 2px solid #fff; }\n        .strip i.sb { background: #f43f5e; opacity: .55; border-bottom: 2px solid #fff; }\n        .tag { white-space: nowrap; display: inline-block; padding: 2px 9px; border-radius: 99px; font-size: 0.74rem; font-weight: 700; }\n        .tag.none { background: rgba(148, 163, 184, 0.2); color: #cbd5e1; }\n        .tag.idle { background: rgba(251, 191, 36, 0.18); color: #fcd34d; }\n        .tag.doing { background: rgba(56, 189, 248, 0.18); color: #7dd3fc; }\n        .tag.done { background: rgba(34, 197, 94, 0.2); color: #86efac; }\n\n        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }\n\n        .q-row { display: grid; grid-template-columns: 70px 1fr 120px 70px; gap: 10px; align-items: center; padding: 6px 0; border-bottom: 1px solid rgba(255, 255, 255, 0.06); font-size: 0.88rem; }\n        .q-row:last-child { border-bottom: none; }\n        .q-row b { color: var(--accent); }\n\n        .chips { display: flex; flex-wrap: wrap; gap: 6px; }\n        .chip { background: rgba(255, 255, 255, 0.08); border: 1px solid var(--border); border-radius: 99px; padding: 3px 11px; font-size: 0.82rem; }\n        .chip small { color: var(--text-muted); margin-left: 4px; }\n\n        /* Thẻ (tab) */\n        .tabs { display: flex; gap: 8px; margin-bottom: 16px; flex-wrap: wrap; }\n        .tab { background: rgba(30, 41, 59, 0.88); border: 1px solid var(--border); color: var(--text-muted); border-radius: 12px; padding: 11px 20px; font-weight: 800; font-size: 0.95rem; cursor: pointer; font-family: inherit; }\n        .tab.active { background: var(--accent); color: #0f172a; border-color: var(--accent); }\n        .tab .badge { background: #f43f5e; color: #fff; border-radius: 99px; padding: 0 7px; font-size: 0.72rem; margin-left: 6px; }\n\n        /* Biểu mẫu & hộp thoại */\n        .form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }\n        .form-grid .full { grid-column: 1 / -1; }\n        .lbl { display: block; font-size: 0.74rem; font-weight: 800; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 4px; }\n        textarea, .form-grid input[type=text], .form-grid select { width: 100%; }\n        textarea { background: rgba(15, 23, 42, 0.9); border: 1px solid var(--border); color: #fff; border-radius: var(--radius-sm); padding: 9px 12px; font-size: 0.92rem; font-family: inherit; outline: none; resize: vertical; }\n        textarea:focus { border-color: var(--primary); }\n        .hint { font-size: 0.8rem; color: var(--text-muted); margin-top: 4px; }\n        .backdrop { position: fixed; inset: 0; background: rgba(2, 6, 23, 0.78); display: flex; align-items: flex-start; justify-content: center; padding: 24px 12px; overflow-y: auto; z-index: 50; }\n        .modal { width: 100%; max-width: 720px; background: #1e293b; border: 1px solid var(--border); border-radius: var(--radius-lg); padding: 22px; box-shadow: var(--shadow); }\n        .modal h3 { font-size: 1.1rem; font-weight: 800; margin-bottom: 14px; }\n        .modal .actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 16px; flex-wrap: wrap; }\n        .btn-danger { background: rgba(244, 63, 94, 0.18); color: #fda4af; border: 1px solid rgba(244, 63, 94, 0.5); }\n        .btn-danger:hover { background: rgba(244, 63, 94, 0.3); }\n        .btn-sm { padding: 5px 10px; font-size: 0.8rem; }\n        .ok-box { background: rgba(34, 197, 94, 0.14); border: 1px solid rgba(34, 197, 94, 0.5); color: #bbf7d0; border-radius: 10px; padding: 10px 12px; font-size: 0.9rem; margin-bottom: 12px; white-space: pre-wrap; }\n        .warn-box { background: rgba(251, 191, 36, 0.12); border: 1px solid rgba(251, 191, 36, 0.5); color: #fde68a; border-radius: 10px; padding: 10px 12px; font-size: 0.9rem; margin-bottom: 12px; white-space: pre-wrap; }\n        td.wrap { white-space: normal; min-width: 220px; max-width: 420px; }\n        #tq-table td, #tq-table th, #tf-table td, #tf-table th { padding: 8px 8px; }\n        #tq-table td, #tf-table td { white-space: normal; }\n        .row-actions { display: flex; gap: 4px; flex-wrap: nowrap; }\n        .blank { color: var(--accent); font-weight: 800; }\n        tr.hid td { opacity: 0.5; }\n        .tag.bad { background: rgba(244, 63, 94, 0.2); color: #fda4af; }\n        .tag.ok { background: rgba(34, 197, 94, 0.2); color: #86efac; }\n        .tag.t-dien_tu { background: rgba(56,189,248,0.16); color: #7dd3fc; }\n        .tag.t-trac_nghiem { background: rgba(167,139,250,0.16); color: #c4b5fd; }\n        .tag.t-viet_lai { background: rgba(244,114,182,0.16); color: #f9a8d4; }\n        .tag.t-word_entry { background: rgba(52,211,153,0.16); color: #6ee7b7; }\n        .opt-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }\n        .opt-grid .optrow { display: flex; align-items: center; gap: 8px; }\n        .opt-grid .optrow b { width: 20px; color: var(--accent); }\n        .tq-filter { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; margin-bottom: 10px; }\n        .tq-filter select, .tq-filter input { background: rgba(15,23,42,0.9); color: var(--text-main); border: 1px solid var(--border); border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; }\n        .tq-filter input { flex: 1; min-width: 180px; }\n        .g-row { display: grid; grid-template-columns: minmax(140px, 250px) 1fr 140px 80px; gap: 10px; align-items: center; padding: 6px 0; border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 0.88rem; }\n        .g-row:last-child { border-bottom: none; }\n        .seg { display: inline-flex; gap: 6px; margin-bottom: 8px; }\n        .seg button.on { border-color: var(--accent); color: var(--accent); }\n        .tq-unit { font-size: 0.78rem; line-height: 1.35; }\n        .tq-unit .u { color: var(--primary); font-weight: 700; }\n        #tq-sel-bar { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; margin-bottom: 10px; padding: 8px 12px; border: 1px dashed var(--accent); border-radius: 10px; background: rgba(251,191,36,0.07); font-size: 0.88rem; }\n        #tq-table input[type=checkbox] { width: 16px; height: 16px; cursor: pointer; }\n        @media (max-width: 700px) { .g-row { grid-template-columns: 1fr 70px; } .g-row .bar, .g-row .muted { display: none; } }\n        .t-row { display: grid; grid-template-columns: 120px 1fr 130px 90px; gap: 10px; align-items: center; padding: 6px 0; border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 0.88rem; }\n        .t-row:last-child { border-bottom: none; }\n        @media (max-width: 700px) { .opt-grid { grid-template-columns: 1fr; } .t-row { grid-template-columns: 100px 1fr 80px; } .t-row .bar { display:none; } }\n        .tag.off { background: rgba(148, 163, 184, 0.2); color: #cbd5e1; }\n        .cls-block { margin-bottom: 14px; }\n        .cls-block summary { cursor: pointer; font-weight: 800; padding: 8px 0; }\n        .stu-row { display: grid; grid-template-columns: 34px 1fr 140px 150px; gap: 8px; align-items: center; padding: 5px 0; border-bottom: 1px solid rgba(255,255,255,.06); font-size: 0.9rem; }\n        .rv-ok { color: #86efac; font-weight: 800; }\n        .rv-bad { color: #fda4af; font-weight: 800; }\n        /* mượt trên điện thoại: nền cố định ở lớp riêng; ô nhập 16px để iPhone không tự phóng to */\n        html { background: #0f172a; }\n        body { background: transparent; background-attachment: scroll; }\n        body::before { content: \"\"; position: fixed; inset: 0; z-index: -1; pointer-events: none; background: radial-gradient(circle at top, #172554 0%, #0f172a 55%, #090d16 100%); }\n        button, .tab { touch-action: manipulation; -webkit-tap-highlight-color: transparent; }\n        @media (max-width: 900px) {\n            input, select, textarea, .field { font-size: 16px !important; }\n        }\n        @media (max-width: 900px) {\n            .form-grid { grid-template-columns: 1fr; }\n            .stu-row { grid-template-columns: 28px 1fr 110px; }\n            .stu-row .del { grid-column: 1 / -1; }\n            .kpi-grid { grid-template-columns: repeat(2, 1fr); }\n            .grid-2 { grid-template-columns: 1fr; }\n        }\n    </style>\n</head>\n<body>\n\n    <!-- ===================== CỔNG MẬT KHẨU GIÁO VIÊN ===================== -->\n    <div id=\"gate\" class=\"gate\">\n        <form class=\"panel gate-card\" onsubmit=\"return submitKey(event)\">\n            <div style=\"text-align:center; margin-bottom: 16px;\">\n                <div class=\"logo\" style=\"margin: 0 auto 10px;\">📊</div>\n                <h1 style=\"font-size: 1.3rem; font-weight: 800;\">Bảng điều khiển giáo viên</h1>\n                <p class=\"muted\">Word Form 9 by NMC</p>\n            </div>\n            <div id=\"gate-error\" class=\"error-box hidden\"></div>\n            <input type=\"password\" id=\"gate-key\" placeholder=\"Mật khẩu giáo viên\" autocomplete=\"off\">\n            <button id=\"gate-btn\" type=\"submit\" class=\"btn btn-accent\" style=\"width: 100%; justify-content: center; padding: 12px;\">Mở bảng điều khiển</button>\n        </form>\n    </div>\n\n    <!-- ===================== BẢNG ĐIỀU KHIỂN ===================== -->\n    <div id=\"app\" class=\"container hidden\">\n        <header class=\"app-header\">\n            <div class=\"brand\">\n                <div class=\"logo\">📊</div>\n                <div>\n                    <h1>Trang giáo viên</h1>\n                    <small id=\"updated-at\">Word Form 9 by NMC</small>\n                </div>\n            </div>\n            <div style=\"display:flex; gap:8px; align-items:center; flex-wrap:wrap;\">\n                <a id=\"student-link\" class=\"btn btn-ghost\" href=\"#\" target=\"_blank\" rel=\"noopener\" title=\"Link để gửi cho học sinh\">🔗 Link học sinh</a>\n                <button class=\"btn btn-ghost\" onclick=\"logoutTeacher()\">🚪 Thoát</button>\n            </div>\n        </header>\n\n        <div id=\"load-error\" class=\"error-box hidden\"></div>\n\n        <nav class=\"tabs\">\n            <button class=\"tab active\" id=\"tab-results\" onclick=\"showTab('results')\">📊 Kết quả</button>\n            <button class=\"tab\" id=\"tab-questions\" onclick=\"showTab('questions')\">📝 Câu hỏi <span id=\"q-badge\" class=\"badge hidden\"></span></button>\n            <button class=\"tab\" id=\"tab-roster\" onclick=\"showTab('roster')\">👥 Lớp học</button>\n            <button class=\"tab\" id=\"tab-links\" onclick=\"showTab('links')\">🔗 Link lớp</button>\n            <button class=\"tab\" id=\"tab-ac\" onclick=\"showTab('ac')\">🛡 Chống gian lận</button>\n        </nav>\n\n        <!-- ============ THẺ 1: KẾT QUẢ ============ -->\n        <div id=\"pane-results\">\n            <div class=\"panel toolbar\">\n                <select id=\"f-khoi\" onchange=\"setKhoi(this.value)\" aria-label=\"Lọc theo khối\"></select>\n                <select id=\"f-class\" onchange=\"setClass(this.value)\"></select>\n                <select id=\"f-status\" onchange=\"view.status = this.value; render();\">\n                    <option value=\"\">Mọi trạng thái</option>\n                    <option value=\"none\">Chưa đăng nhập</option>\n                    <option value=\"idle\">Đã đăng nhập, chưa làm câu nào</option>\n                    <option value=\"doing\">Đang làm</option>\n                    <option value=\"done\">Đã làm hết các câu</option>\n                </select>\n                <input type=\"text\" id=\"f-search\" class=\"grow\" placeholder=\"🔍 Tìm theo họ tên...\" oninput=\"view.q = this.value; render();\">\n                <span class=\"muted\" id=\"count-label\"></span>\n                <label class=\"muted\" style=\"display:flex; align-items:center; gap:6px; cursor:pointer;\">\n                    <input type=\"checkbox\" id=\"auto-refresh\" checked> Tự làm mới 60 giây\n                </label>\n                <button class=\"btn btn-accent\" id=\"refresh-btn\" onclick=\"loadData()\">🔄 Làm mới</button>\n                <button class=\"btn btn-ghost\" onclick=\"exportCsv()\">⬇️ Xuất CSV tổng hợp</button>\n                <button class=\"btn btn-ghost\" id=\"export-detail-btn\" onclick=\"exportDetailCsv()\">⬇️ Xuất CSV từng câu</button>\n                <button class=\"btn btn-ghost\" onclick=\"exportEachClass()\" title=\"Tải riêng một file cho mỗi lớp đang hiển thị\">⬇️ Xuất mỗi lớp một file</button>\n            </div>\n\n            <div class=\"kpi-grid\" id=\"kpis\"></div>\n\n            <div class=\"panel\" id=\"class-panel\">\n                <h2>Tổng hợp theo lớp</h2>\n                <div class=\"table-wrap\"><table id=\"class-table\"></table></div>\n                <p class=\"muted\" style=\"margin-top:8px;\">Bấm vào một lớp để xem chi tiết từng học sinh.</p>\n            </div>\n\n            <div class=\"panel\">\n                <h2>Chi tiết từng học sinh</h2>\n                <div class=\"table-wrap\"><table id=\"student-table\"></table></div>\n                <p class=\"muted\" style=\"margin-top:8px;\">\n                    Mỗi ô nhỏ là kết quả mới nhất của một câu:\n                    <span class=\"strip\"><i class=\"c\"></i></span> đúng\n                    <span class=\"strip\"><i class=\"w\"></i></span> sai\n                    <span class=\"strip\"><i class=\"r\"></i></span> xem đáp án\n                    <span class=\"strip\"><i></i></span> chưa làm. Rê chuột lên ô để xem từ gốc / dạng bài.\n                    Bấm <b>tên học sinh</b> để xem bài làm từng câu.\n                    <br>Điểm /10 = số câu đúng ÷ số câu được giao cho lớp đó × 10 (xem đáp án tính là sai).\n                </p>\n            </div>\n\n            <div class=\"panel hidden\" id=\"type-panel\">\n                <h2>Tỉ lệ đúng theo dạng bài</h2>\n                <div id=\"type-list\"></div>\n                <p class=\"muted\" style=\"margin-top:8px;\">Tính trên các câu học sinh đã làm (xem đáp án tính là sai), theo lớp đang chọn.</p>\n            </div>\n\n            <div class=\"panel hidden\" id=\"group-panel\">\n                <h2>Tỉ lệ đúng theo Unit / chủ đề</h2>\n                <div class=\"seg\">\n                    <button class=\"btn btn-ghost btn-sm on\" id=\"grp-unit\" onclick=\"setGroupBy('unit')\">Theo Unit</button>\n                    <button class=\"btn btn-ghost btn-sm\" id=\"grp-topic\" onclick=\"setGroupBy('topic')\">Theo chủ đề</button>\n                </div>\n                <div id=\"group-list\"></div>\n                <p class=\"muted\" style=\"margin-top:8px;\">⚠ = tỉ lệ đúng dưới 50% (phần cần ôn thêm). Tính theo lớp đang chọn; Unit/chủ đề gán ở thẻ Câu hỏi.</p>\n            </div>\n\n            <div class=\"grid-2\">\n                <div class=\"panel\">\n                    <h2>Câu cả lớp sai nhiều nhất</h2>\n                    <div id=\"q-list\"></div>\n                    <button class=\"btn btn-ghost\" id=\"q-toggle\" style=\"margin-top:10px;\" onclick=\"view.allQ = !view.allQ; renderQuestions();\"></button>\n                </div>\n                <div class=\"panel\">\n                    <h2>Cần nhắc nhở</h2>\n                    <div style=\"margin-bottom:14px;\">\n                        <div class=\"muted\" style=\"margin-bottom:6px;\">Chưa đăng nhập (<span id=\"n-none\">0</span>)</div>\n                        <div class=\"chips\" id=\"list-none\"></div>\n                    </div>\n                    <div>\n                        <div class=\"muted\" style=\"margin-bottom:6px;\">Đã đăng nhập nhưng chưa làm câu nào (<span id=\"n-idle\">0</span>)</div>\n                        <div class=\"chips\" id=\"list-idle\"></div>\n                    </div>\n                </div>\n            </div>\n        </div>\n\n        <!-- ============ THẺ: LINK TỪNG LỚP / KHỐI ============ -->\n        <div id=\"pane-links\" class=\"hidden\">\n            <div class=\"panel\">\n                <div style=\"display:flex; flex-wrap:wrap; gap:8px; align-items:center; justify-content:space-between; margin-bottom:8px;\">\n                    <h2 style=\"margin:0;\">Link gửi học sinh theo lớp / khối</h2>\n                    <button class=\"btn btn-accent\" onclick=\"copyAllLinks()\">📋 Sao chép tất cả (tên lớp + link)</button>\n                </div>\n                <div id=\"links-msg\"></div>\n                <p class=\"muted\" style=\"margin-bottom:10px;\">Gửi link đúng lớp cho học sinh: em mở link là chỉ thấy lớp của mình (ô chọn lớp được chọn sẵn) và chỉ làm các câu được giao cho lớp / khối đó.\n                    Link khối cho phép chọn trong các lớp của khối. Link \"cả trường\" cho chọn mọi lớp.</p>\n                <div class=\"table-wrap\"><table id=\"links-table\"></table></div>\n            </div>\n        </div>\n\n        <!-- ============ THẺ: CHỐNG GIAN LẬN (RỜI TRANG) ============ -->\n        <div id=\"pane-ac\" class=\"hidden\">\n            <div class=\"panel\">\n                <h2>Chống gian lận: theo dõi học sinh rời trang</h2>\n                <p class=\"muted\" style=\"margin-bottom:10px;\">Chọn cách học sinh làm bài: <b>tự do</b> hoặc <b>bị nhắc nhở</b> khi rời trang (đổi tab, thu nhỏ trình duyệt, chuyển sang ứng dụng khác, tắt màn hình điện thoại).\n                    Có thể đặt khác nhau cho từng lớp. Học sinh đang làm đề sẽ thấy thông báo và số lần đã rời trang. Rời trang quá số lần cho phép thì bài <b>tự nộp</b>.\n                    Hệ thống chỉ phát hiện và ghi lại; không chặn được chụp màn hình hay dùng thiết bị khác, nên hãy xem số lần rời trang như tham khảo.</p>\n                <div id=\"ac-msg\"></div>\n                <div class=\"form-grid\" style=\"max-width:760px; margin-bottom:12px;\">\n                    <div class=\"full\"><label class=\"lbl\" for=\"ac-mode\">Cách làm bài mặc định</label><select id=\"ac-mode\"></select>\n                        <div class=\"hint\"><b>Tự do</b>: không theo dõi (như trước). <b>Nhắc nhở khi Làm đề tổng hợp</b>: đề có giám sát; rời trang thì nhắc và ghi lại, quá giới hạn thì tự nộp, đồng thời chặn sao chép / dán trong đề. <b>Nhắc nhở cả khi luyện tập</b>: như trên, và luyện tập từng câu / sổ tay cũng được ghi nhận (chỉ nhắc, không tự nộp).</div></div>\n                    <div><label class=\"lbl\" for=\"ac-limit\">Số lần rời trang được phép (1–20)</label><input type=\"text\" id=\"ac-limit\" inputmode=\"numeric\" placeholder=\"4\"></div>\n                </div>\n                <h3 style=\"margin:14px 0 6px;\">Cài đặt riêng từng lớp (không bắt buộc)</h3>\n                <div class=\"table-wrap\"><table id=\"ac-class-table\"></table></div>\n                <div style=\"margin-top:12px;\"><button class=\"btn btn-accent\" onclick=\"saveAc()\">💾 Lưu cài đặt</button></div>\n            </div>\n            <div class=\"panel\">\n                <div style=\"display:flex; flex-wrap:wrap; gap:8px; align-items:center; justify-content:space-between; margin-bottom:8px;\">\n                    <h2 style=\"margin:0;\">Nhật ký rời trang</h2>\n                    <div style=\"display:flex; gap:8px;\">\n                        <button class=\"btn btn-ghost\" onclick=\"loadAc()\">🔄 Làm mới</button>\n                        <button class=\"btn btn-ghost\" onclick=\"exportAcCsv()\">⬇️ Xuất CSV nhật ký</button>\n                    </div>\n                </div>\n                <div class=\"tq-filter\">\n                    <select id=\"ac-f-class\" onchange=\"renderAcLog()\" aria-label=\"Lọc theo lớp\"></select>\n                    <label class=\"muted\">Từ ngày <input type=\"date\" id=\"ac-f-from\" onchange=\"renderAcLog()\"></label>\n                    <label class=\"muted\">đến <input type=\"date\" id=\"ac-f-to\" onchange=\"renderAcLog()\"></label>\n                </div>\n                <h3 style=\"margin:10px 0 6px;\">Tổng hợp theo học sinh</h3>\n                <div class=\"table-wrap\"><table id=\"ac-sum-table\"></table></div>\n                <h3 style=\"margin:14px 0 6px;\">Các lần rời trang gần đây</h3>\n                <div class=\"table-wrap\" style=\"max-height:50vh; overflow:auto;\"><table id=\"ac-log-table\"></table></div>\n                <p class=\"muted\" style=\"margin-top:8px;\">Mỗi dòng là một lần học sinh quay lại trang sau khi rời đi. Dữ liệu lưu ở sheet <b>GiamSat</b> trong Google Sheet (bạn có thể xóa dòng ở đó để bắt đầu đợt mới). Học sinh có thể bị ghi nhận nhầm khi có cuộc gọi hoặc thông báo hệ thống, nên hãy xem thời gian rời và hỏi lại em trước khi kết luận.</p>\n            </div>\n        </div>\n\n        <!-- ============ THẺ 2: CÂU HỎI ============ -->\n        <div id=\"pane-questions\" class=\"hidden\">\n            <div class=\"panel\">\n                <div style=\"display:flex; flex-wrap:wrap; gap:8px; align-items:center; justify-content:space-between; margin-bottom:12px;\">\n                    <h2 style=\"margin:0;\">Câu hỏi <span class=\"muted\" id=\"tq-count\"></span></h2>\n                    <div style=\"display:flex; gap:8px; flex-wrap:wrap;\">\n                        <button class=\"btn btn-accent\" onclick=\"openQForm(0)\">➕ Thêm câu</button>\n                        <button class=\"btn btn-ghost\" onclick=\"openBulkQ()\">📋 Dán nhiều câu</button>\n                        <button class=\"btn btn-ghost\" onclick=\"exportQuestionsCsv()\" title=\"Tải các câu đang lọc (theo khối/lớp, Unit...) thành file để gửi hoặc nhập lại\">⬇️ Xuất câu đang lọc</button>\n                        <button class=\"btn btn-ghost\" id=\"regrade-btn\" onclick=\"regradeAll()\">🔁 Chấm lại bài đã nộp</button>\n                        <button class=\"btn btn-ghost hidden\" id=\"seed-btn\" onclick=\"seedQuestions()\">📦 Nạp câu mẫu (76 câu)</button>\n                    </div>\n                </div>\n                <div id=\"tq-msg\"></div>\n                <div class=\"tq-filter\">\n                    <select id=\"tq-type\" onchange=\"view.qType = this.value; view.qLimit = 100; renderQuestionsTab();\" aria-label=\"Lọc theo dạng bài\"></select>\n                    <select id=\"tq-scope\" onchange=\"view.qScope = this.value; view.qLimit = 100; renderQuestionsTab();\" aria-label=\"Lọc theo khối/lớp\"></select>\n                    <select id=\"tq-unit\" onchange=\"view.qUnit = this.value; view.qTopic = 'all'; view.qLimit = 100; renderQuestionsTab();\" aria-label=\"Lọc theo Unit\"></select>\n                    <select id=\"tq-topic\" onchange=\"view.qTopic = this.value; view.qLimit = 100; renderQuestionsTab();\" aria-label=\"Lọc theo chủ đề\"></select>\n                    <input type=\"text\" id=\"tq-search\" placeholder=\"🔍 Tìm trong câu hỏi, đáp án, Unit, chủ đề...\" oninput=\"view.qSearch = this.value; view.qLimit = 100; renderQuestionsTab();\">\n                </div>\n                <div id=\"tq-sel-bar\" class=\"hidden\">\n                    <span id=\"tq-sel-text\"></span>\n                    <button class=\"btn btn-accent btn-sm\" onclick=\"openScopeForm()\">👥 Gán khối / lớp</button>\n                    <button class=\"btn btn-accent btn-sm\" onclick=\"openTagForm()\">🏷️ Gán Unit / Chủ đề</button>\n                    <button class=\"btn btn-ghost btn-sm\" onclick=\"clearSel()\">Bỏ chọn</button>\n                </div>\n                <div class=\"table-wrap\"><table id=\"tq-table\"></table></div>\n                <p class=\"muted\" style=\"margin-top:8px;\">\n                    Học sinh thấy ngay câu mới khi tải lại trang. Nếu sửa <b>đáp án</b> của câu đã có học sinh làm, bấm\n                    <b>Chấm lại bài đã nộp</b> để điểm của mọi học sinh được tính theo đáp án mới.\n                    Câu bị <span class=\"tag bad\">⚠</span> sẽ chưa hiện cho học sinh cho đến khi sửa xong.\n                </p>\n            </div>\n\n            <div class=\"panel\">\n                <div style=\"display:flex; flex-wrap:wrap; gap:8px; align-items:center; justify-content:space-between; margin-bottom:12px;\">\n                    <h2 style=\"margin:0;\">Mục từ điển (cho dạng Word entry) <span class=\"muted\" id=\"te-count\"></span></h2>\n                    <button class=\"btn btn-ghost\" onclick=\"openEForm('')\">➕ Thêm mục từ</button>\n                </div>\n                <div id=\"te-msg\"></div>\n                <div class=\"table-wrap\"><table id=\"te-table\"></table></div>\n                <p class=\"muted\" style=\"margin-top:8px;\">Mục từ gồm: từ cần tra, phiên âm, từ loại, nghĩa và 3–5 câu ví dụ. Cụm 2–3 từ học sinh cần học viết giữa hai dấu <b>**</b> (ví dụ <b>**make a decision**</b>). Đáp án của câu Word entry phải là một cụm in đậm đó.</p>\n            </div>\n\n            <div class=\"panel\">\n                <div style=\"display:flex; flex-wrap:wrap; gap:8px; align-items:center; justify-content:space-between; margin-bottom:12px;\">\n                    <h2 style=\"margin:0;\">Gia đình từ (Sổ tay) <span class=\"muted\" id=\"tf-count\"></span></h2>\n                    <button class=\"btn btn-ghost\" onclick=\"openFForm('')\">➕ Thêm gia đình từ</button>\n                </div>\n                <div class=\"table-wrap\"><table id=\"tf-table\"></table></div>\n            </div>\n        </div>\n\n        <!-- ============ THẺ 3: LỚP HỌC ============ -->\n        <div id=\"pane-roster\" class=\"hidden\">\n            <div class=\"panel\">\n                <h2>Thêm học sinh vào lớp</h2>\n                <div id=\"tr-msg\"></div>\n                <div class=\"form-grid\">\n                    <div>\n                        <label class=\"lbl\" for=\"tr-cls\">Tên lớp</label>\n                        <input type=\"text\" id=\"tr-cls\" list=\"tr-cls-list\" placeholder=\"Ví dụ 9A1\">\n                        <datalist id=\"tr-cls-list\"></datalist>\n                        <div class=\"hint\">Để trống nếu dán 2 cột <i>Lớp ⇥ Họ tên</i> (copy từ Excel).</div>\n                    </div>\n                    <div></div>\n                    <div class=\"full\">\n                        <label class=\"lbl\" for=\"tr-text\">Danh sách họ tên (mỗi dòng một em)</label>\n                        <textarea id=\"tr-text\" rows=\"9\" placeholder=\"Nguyễn Văn An&#10;Lê Thị Hoa&#10;...\"></textarea>\n                        <div class=\"hint\">Số thứ tự đầu dòng tự bỏ; chữ IN HOA hoặc chữ thường toàn bộ tự đổi thành Viết Hoa Đầu Từ. Em nào đã có trong lớp sẽ được bỏ qua.</div>\n                    </div>\n                </div>\n                <div class=\"actions\" style=\"margin-top:12px;\"><button class=\"btn btn-accent\" id=\"tr-btn\" onclick=\"addRoster()\">Thêm vào danh sách</button></div>\n            </div>\n            <div class=\"panel\">\n                <h2>Danh sách hiện có <span class=\"muted\" id=\"tr-count\"></span></h2>\n                <div id=\"tr-list\"></div>\n                <p class=\"muted\" style=\"margin-top:8px;\">\n                    <b>Mã học sinh</b> (không bắt buộc): nếu đặt mã, học sinh phải nhập thêm mã này mới vào được — dùng khi sợ bạn khác gõ tên hộ.\n                    Xóa học sinh khỏi danh sách thì em đó không vào làm bài được nữa (kết quả cũ vẫn còn trong Google Sheet).\n                </p>\n            </div>\n        </div>\n    </div>\n\n    <!-- Hộp thoại dùng chung -->\n    <div id=\"modal\" class=\"backdrop hidden\"><div class=\"modal\" id=\"modal-body\"></div></div>\n\n    <script>\n        // ==========================================================================\n        // Dán URL Web App của Google Apps Script (kết thúc bằng /exec) vào dòng dưới đây.\n        // ==========================================================================\n        const API_URL = \"DÁN_URL_WEB_APP_VÀO_ĐÂY\";\n        const KEY_STORE = \"wf9_teacher_key\";\n\n        let teacherKey = \"\";\n        let DATA = null, DETAIL = null, TQ = null, TR = null, modalAction = null;\n        let timer = null;\n        const view = { cls: \"\", status: \"\", q: \"\", sortKey: \"cls\", sortDir: 1, allQ: false, qType: \"all\", qSearch: \"\", qUnit: \"all\", qTopic: \"all\", qLimit: 100, sel: {}, groupBy: \"unit\", khoi: \"\", qScope: \"all\" };\n\n        // ---------------------------------------------------------------- tiện ích\n\n        // ---------------------------------------------------------------- v5: dạng bài\n        const TYPE_LABEL = { dien_tu: \"Điền từ\", trac_nghiem: \"Trắc nghiệm\", viet_lai: \"Viết lại câu\", word_entry: \"Word entry\" };\n        const TYPE_ORDER = [\"dien_tu\", \"trac_nghiem\", \"viet_lai\", \"word_entry\"];\n        function qTypeOf(q) { return TYPE_LABEL[q && q.type] ? q.type : \"dien_tu\"; }\n        function typeTag(q) { const t = qTypeOf(q); return `<span class=\"tag t-${t}\">${TYPE_LABEL[t]}</span>`; }\n        function qLabelOf(q) {\n            const t = qTypeOf(q);\n            if (t === \"dien_tu\") return q.root;\n            if (t === \"word_entry\") return \"📖 \" + (q.entry || \"\");\n            const sn = String(q.sentence || \"\").replace(/\\*\\*/g, \"\");\n            return sn.length > 46 ? sn.slice(0, 46) + \"…\" : sn;\n        }\n        function shortAns(a) { a = String(a || \"\").replace(/\\*\\*/g, \"\"); return a.length > 56 ? a.slice(0, 56) + \"…\" : a; }\n        function ansFull(q) {\n            if (q.answerFull) return q.answerFull;\n            if (qTypeOf(q) === \"trac_nghiem\" && q.opts) { const i = \"ABCD\".indexOf((q.answer || \"\").toUpperCase()); return i >= 0 ? q.answer + \". \" + (q.opts[i] || \"\") : q.answer; }\n            return q.answer;\n        }\n        function richHtml(s) { return esc(s).replace(/\\*\\*(.+?)\\*\\*/g, \"<b>$1</b>\"); }\n        function rootCell(q) {\n            const t = qTypeOf(q);\n            if (t === \"word_entry\") return q.entry ? \"📖 \" + esc(q.entry) : '<span class=\"muted\">—</span>';\n            return esc(q.root) || '<span class=\"muted\">—</span>';\n        }\n\n        function esc(s) {\n            return String(s == null ? \"\" : s).replace(/&/g, \"&amp;\").replace(/</g, \"&lt;\")\n                .replace(/>/g, \"&gt;\").replace(/\"/g, \"&quot;\").replace(/'/g, \"&#39;\");\n        }\n\n        function fold(s) {\n            return String(s || \"\").normalize(\"NFD\").replace(/[̀-ͯ]/g, \"\").replace(/đ/g, \"d\").toLowerCase();\n        }\n\n        function natCmp(a, b) {\n            return String(a).localeCompare(String(b), \"vi\", { numeric: true });\n        }\n\n        function total() { return (DATA && DATA.total) || 0; }\n\n        // ---------------------------------------------------------------- v7: khối / lớp\n        function khoiOf(cls) { const m = String(cls || \"\").match(/\\d+/); return m ? String(parseInt(m[0], 10)) : \"\"; }\n        function classKey(s) { return fold(s).replace(/^(lop|class)\\s*/, \"\").replace(/[\\s.\\-_]+/g, \"\"); }\n        function scopeTokens(sc) { return String(sc || \"\").split(/[|,;\\n]+/).map(t => t.trim()).filter(Boolean); }\n        function scopeInfo(t) {\n            const raw = fold(t).replace(/\\s+/g, \"\");\n            if ([\"tatca\", \"all\", \"chung\", \"moilop\", \"tatcacaclop\", \"moikhoi\"].includes(raw)) return { kind: \"all\", key: \"all\" };\n            const m = /^(?:khoi|k)?(\\d{1,2})$/.exec(raw);\n            if (m) return { kind: \"khoi\", key: String(parseInt(m[1], 10)) };\n            return { kind: \"lop\", key: classKey(t) };\n        }\n        // câu có \"Áp dụng cho\" = sc thì lớp cls được giao không (trống = mọi lớp)\n        function inScope(sc, cls) {\n            const toks = scopeTokens(sc);\n            if (!toks.length) return true;\n            const ck = classKey(cls), kh = khoiOf(cls);\n            return toks.some(t => { const i = scopeInfo(t); return i.kind === \"all\" || (i.kind === \"khoi\" ? (kh && i.key === kh) : i.key === ck); });\n        }\n        function nOf(s) { return s.n != null ? s.n : total(); }\n        function isAsg(s, i) { return !s.asg || s.asg[i] !== \"0\"; }\n        function allKhoi() {\n            const set = {};\n            (DATA ? DATA.classes : []).forEach(c => { const k = khoiOf(c); if (k) set[k] = 1; });\n            return Object.keys(set).sort((a, b) => a - b);\n        }\n        function selClasses() { return DATA.classes.filter(c => (!view.cls || c === view.cls) && (!view.khoi || khoiOf(c) === view.khoi)); }\n        function safeName(x) { return String(x).replace(/[^\\w-]+/g, \"_\").replace(/^_+|_+$/g, \"\"); }\n        const isGood = v => v === \"Đ\" || v === \"T+\";\n        const isBad = v => v === \"S\" || v === \"X\" || v === \"T-\";\n        const el = id => document.getElementById(id);\n\n        function pct(x) { return Math.round(x * 100) + \"%\"; }\n\n        function ago(iso) {\n            if (!iso) return \"—\";\n            const t = new Date(iso).getTime();\n            if (isNaN(t)) return \"—\";\n            const m = Math.round((Date.now() - t) / 60000);\n            if (m < 1) return \"vừa xong\";\n            if (m < 60) return m + \" phút trước\";\n            const h = Math.round(m / 60);\n            if (h < 24) return h + \" giờ trước\";\n            const d = Math.round(h / 24);\n            return d + \" ngày trước\";\n        }\n\n        function statusOf(s) {\n            if (!s.loggedIn) return \"none\";\n            if (s.done === 0) return \"idle\";\n            if (nOf(s) > 0 && s.done >= nOf(s)) return \"done\";\n            return \"doing\";\n        }\n\n        const STATUS_LABEL = { none: \"Chưa đăng nhập\", idle: \"Chưa làm\", doing: \"Đang làm\", done: \"Hoàn thành\" };\n\n        function score1(x) { return (Math.round(x * 10) / 10).toFixed(1); }\n\n        // ---------------------------------------------------------------- gọi API\n        async function apiPost(payload) {\n            const ctrl = new AbortController();\n            const t = setTimeout(() => ctrl.abort(), 30000);\n            try {\n                const res = await fetch(API_URL, {\n                    method: \"POST\", redirect: \"follow\", signal: ctrl.signal,\n                    headers: { \"Content-Type\": \"text/plain;charset=utf-8\" },\n                    body: JSON.stringify(payload)\n                });\n                return await res.json();\n            } finally { clearTimeout(t); }\n        }\n\n        function apiConfigured() {\n            return /^https:\\/\\/script\\.google\\.com\\/.+\\/exec$/.test(API_URL);\n        }\n\n        // ---------------------------------------------------------------- đăng nhập giáo viên\n        function showGateError(msg) {\n            const el = document.getElementById(\"gate-error\");\n            el.textContent = msg || \"\";\n            el.classList.toggle(\"hidden\", !msg);\n        }\n\n        async function submitKey(ev) {\n            if (ev && ev.preventDefault) ev.preventDefault();\n            const key = document.getElementById(\"gate-key\").value;\n            if (!key) { showGateError(\"Hãy nhập mật khẩu giáo viên.\"); return false; }\n            const btn = document.getElementById(\"gate-btn\");\n            btn.disabled = true; btn.textContent = \"Đang kiểm tra...\";\n            showGateError(\"\");\n            teacherKey = key;\n            const ok = await loadData(true);\n            btn.disabled = false; btn.textContent = \"Mở bảng điều khiển\";\n            if (ok) {\n                try { sessionStorage.setItem(KEY_STORE, key); } catch (e) {}\n                enterApp();\n            }\n            return false;\n        }\n\n        function logoutTeacher() {\n            try { sessionStorage.removeItem(KEY_STORE); } catch (e) {}\n            teacherKey = \"\"; DATA = null;\n            if (timer) clearInterval(timer);\n            location.reload();\n        }\n\n        // ---------------------------------------------------------------- tải dữ liệu\n        async function loadData(fromGate) {\n            const btn = document.getElementById(\"refresh-btn\");\n            if (btn) { btn.disabled = true; }\n            try {\n                const r = await apiPost({ action: \"dashboard\", key: teacherKey });\n                if (!r || !r.ok) {\n                    if (r && r.error === \"BAD_KEY\") {\n                        if (fromGate) showGateError(\"Mật khẩu chưa đúng.\");\n                        else logoutTeacher();\n                    } else {\n                        const msg = \"Máy chủ báo lỗi\" + (r && r.message ? \": \" + r.message : \".\");\n                        if (fromGate) showGateError(msg); else showLoadError(msg);\n                    }\n                    return false;\n                }\n                DATA = r; DETAIL = null;\n                showLoadError(\"\");\n                buildClassSelect();\n                render();\n                const when = new Date(r.generated);\n                document.getElementById(\"updated-at\").textContent =\n                    \"Cập nhật lúc \" + (isNaN(when) ? r.generated : when.toLocaleTimeString(\"vi-VN\") + \" \" + when.toLocaleDateString(\"vi-VN\"));\n                return true;\n            } catch (e) {\n                const msg = \"Không kết nối được máy chủ. Kiểm tra mạng hoặc URL Web App.\";\n                if (fromGate) showGateError(msg); else showLoadError(msg);\n                return false;\n            } finally {\n                if (btn) btn.disabled = false;\n            }\n        }\n\n        function showLoadError(msg) {\n            const el = document.getElementById(\"load-error\");\n            el.textContent = msg || \"\";\n            el.classList.toggle(\"hidden\", !msg);\n        }\n\n        function startAutoRefresh() {\n            if (timer) clearInterval(timer);\n            timer = setInterval(() => {\n                if (document.getElementById(\"auto-refresh\").checked && !document.hidden) loadData();\n            }, 60000);\n        }\n\n        // ---------------------------------------------------------------- bộ lọc\n        function buildClassSelect() {\n            const ks = document.getElementById(\"f-khoi\");\n            const kl = allKhoi();\n            ks.innerHTML = '<option value=\"\">Tất cả các khối</option>' + kl.map(k => `<option value=\"${k}\">Khối ${k}</option>`).join(\"\");\n            if (!kl.includes(view.khoi)) view.khoi = \"\";\n            ks.value = view.khoi;\n            ks.classList.toggle(\"hidden\", kl.length < 2);\n            const sel = document.getElementById(\"f-class\");\n            const cur = view.cls;\n            const list = DATA.classes.filter(c => !view.khoi || khoiOf(c) === view.khoi);\n            sel.innerHTML = '<option value=\"\">Tất cả các lớp' + (view.khoi ? \" của khối \" + view.khoi : \"\") + '</option>' +\n                list.map(c => `<option value=\"${esc(c)}\">Lớp ${esc(c)}</option>`).join(\"\");\n            sel.value = list.includes(cur) ? cur : \"\";\n            view.cls = sel.value;\n        }\n\n        function setKhoi(k) { view.khoi = k; view.cls = \"\"; buildClassSelect(); render(); }\n        function setClass(c) { view.cls = c; document.getElementById(\"f-class\").value = c; render(); }\n\n        function filtered() {\n            const q = fold(view.q.trim());\n            return DATA.students.filter(s =>\n                (!view.khoi || khoiOf(s.cls) === view.khoi) &&\n                (!view.cls || s.cls === view.cls) &&\n                (!view.status || statusOf(s) === view.status) &&\n                (!q || fold(s.name).includes(q))\n            );\n        }\n\n        function inClass() {\n            return DATA.students.filter(s => (!view.khoi || khoiOf(s.cls) === view.khoi) && (!view.cls || s.cls === view.cls));\n        }\n\n        // ---------------------------------------------------------------- hiển thị\n        function render() {\n            if (!DATA) return;\n            const list = filtered();\n            document.getElementById(\"count-label\").textContent = list.length + \" / \" + DATA.students.length + \" học sinh\";\n            renderKpis();\n            renderClassTable();\n            renderStudents(list);\n            renderQuestions();\n            renderTypeStats();\n            renderGroupStats();\n            renderMissing();\n        }\n\n        function renderKpis() {\n            const scope = inClass();\n            const logged = scope.filter(s => s.loggedIn).length;\n            const started = scope.filter(s => s.done > 0);\n            const finished = scope.filter(s => nOf(s) > 0 && s.done >= nOf(s)).length;\n            const avgScore = started.length ? started.reduce((a, s) => a + s.score, 0) / started.length : 0;\n            const ns = {}; scope.forEach(s => { ns[nOf(s)] = 1; });\n            const nk = Object.keys(ns);\n            const finLabel = nk.length === 1 ? \"Đã làm hết \" + nk[0] + \" câu\" : \"Đã làm hết bộ câu được giao\";\n            const where = view.cls ? \"lớp \" + view.cls : view.khoi ? \"khối \" + view.khoi : \"toàn trường\";\n            const cards = [\n                [\"Sĩ số\", scope.length, where],\n                [\"Đã đăng nhập\", logged, scope.length ? pct(logged / scope.length) : \"—\"],\n                [\"Đã làm ≥ 1 câu\", started.length, scope.length ? pct(started.length / scope.length) : \"—\"],\n                [finLabel, finished, scope.length ? pct(finished / scope.length) : \"—\"],\n                [\"Điểm TB (thang 10)\", started.length ? score1(avgScore) : \"—\", \"trên các em đã làm\"]\n            ];\n            document.getElementById(\"kpis\").innerHTML = cards.map(c =>\n                `<div class=\"kpi\"><div class=\"label\">${c[0]}</div><div class=\"val\">${c[1]}</div><div class=\"sub\">${c[2]}</div></div>`\n            ).join(\"\");\n        }\n\n        function renderClassTable() {\n            const panel = document.getElementById(\"class-panel\");\n            panel.classList.toggle(\"hidden\", !!view.cls);\n            if (view.cls) return;\n            const rows = selClasses().map(c => {\n                const sc = DATA.students.filter(s => s.cls === c);\n                const started = sc.filter(s => s.done > 0);\n                const avgProg = sc.length ? sc.reduce((a, s) => a + (nOf(s) ? s.done / nOf(s) : 0), 0) / sc.length : 0;\n                const avgSc = started.length ? started.reduce((a, s) => a + s.score, 0) / started.length : 0;\n                const nq = sc.length ? nOf(sc[0]) : 0;\n                return `<tr class=\"clickable\" data-cls=\"${esc(c)}\" onclick=\"setClass(this.dataset.cls)\">\n                    <td><b>${esc(c)}</b></td>\n                    <td class=\"num\">${sc.length}</td>\n                    <td class=\"num\">${nq}</td>\n                    <td class=\"num\">${sc.filter(s => s.loggedIn).length}</td>\n                    <td class=\"num\">${started.length}</td>\n                    <td class=\"num\">${sc.filter(s => nOf(s) > 0 && s.done >= nOf(s)).length}</td>\n                    <td><span class=\"bar\"><i style=\"width:${Math.round(avgProg * 100)}%\"></i></span> ${pct(avgProg)}</td>\n                    <td class=\"num\"><b>${started.length ? score1(avgSc) : \"—\"}</b></td>\n                </tr>`;\n            }).join(\"\");\n            document.getElementById(\"class-table\").innerHTML =\n                `<thead><tr><th>Lớp</th><th>Sĩ số</th><th>Số câu giao</th><th>Đã đăng nhập</th><th>Đã làm</th><th>Hoàn thành</th><th>Tiến độ TB</th><th>Điểm TB</th></tr></thead><tbody>${rows}</tbody>`;\n        }\n\n        const COLUMNS = [\n            { key: \"cls\", label: \"Lớp\" },\n            { key: \"name\", label: \"Họ và tên\" },\n            { key: \"status\", label: \"Trạng thái\" },\n            { key: \"done\", label: \"Tiến độ\" },\n            { key: \"correct\", label: \"Đúng\" },\n            { key: \"wrong\", label: \"Sai\" },\n            { key: \"score\", label: \"Điểm /10\" },\n            { key: \"attempts\", label: \"Lượt\" },\n            { key: \"lv\", label: \"Rời trang\" },\n            { key: \"last\", label: \"Lần cuối\" },\n            { key: \"strip\", label: \"Từng câu\", nosort: true }\n        ];\n\n        function fmtSecs(n) { n = Math.round(n || 0); return n < 60 ? n + \" giây\" : Math.floor(n / 60) + \" phút \" + (n % 60) + \" giây\"; }\n        function lvCell(s) {\n            if (!s.lv) return s.acm && s.acm !== \"free\" ? '<span class=\"muted\">0</span>' : '<span class=\"muted\">—</span>';\n            return `<span class=\"tag ${s.aus ? \"bad\" : \"idle\"}\" title=\"${s.lv} lần rời trang, tổng ${esc(fmtSecs(s.lsec))}${s.aus ? \", \" + s.aus + \" bài tự nộp\" : \"\"}\">${s.lv} lần · ${esc(fmtSecs(s.lsec))}${s.aus ? \" · tự nộp\" : \"\"}</span>`;\n        }\n\n        function sortValue(s, key) {\n            switch (key) {\n                case \"status\": return [\"none\", \"idle\", \"doing\", \"done\"].indexOf(statusOf(s));\n                case \"last\": return s.last ? new Date(s.last).getTime() : 0;\n                default: return s[key];\n            }\n        }\n\n        function sortBy(key) {\n            if (view.sortKey === key) view.sortDir = -view.sortDir;\n            else { view.sortKey = key; view.sortDir = (key === \"cls\" || key === \"name\") ? 1 : -1; }\n            render();\n        }\n\n        const STRIP = { \"Đ\": [\"c\", \"đúng\"], \"S\": [\"w\", \"sai\"], \"X\": [\"r\", \"xem đáp án\"] };\n\n        // Ngân hàng lớn (> 200 câu): vẽ dải kết quả bằng canvas gộp theo nhóm câu, tránh hàng chục nghìn ô HTML\n        const STRIP_BIG = 200, STRIP_BUCKETS = 150;\n        let STRIP_LIST = [];\n        function stripBuckets(s) {\n            const n = total(), k = Math.ceil(n / STRIP_BUCKETS), out = [];\n            for (let b = 0; b * k < n; b++) {\n                const o = { c: 0, w: 0, r: 0, u: 0, from: b * k, to: Math.min(n, (b + 1) * k) - 1 };\n                for (let i = o.from; i <= o.to; i++) {\n                    if (!isAsg(s, i)) continue;\n                    const ch = s.states[i] || \"\";\n                    if (ch === \"Đ\") o.c++; else if (ch === \"S\") o.w++; else if (ch === \"X\") o.r++; else o.u++;\n                }\n                out.push(o);\n            }\n            return out;\n        }\n        function drawStrips() {\n            const dpr = Math.min(2, window.devicePixelRatio || 1);\n            document.querySelectorAll(\"canvas.strip-cv\").forEach(cv => {\n                const s = STRIP_LIST[+cv.dataset.i]; if (!s) return;\n                const bs = stripBuckets(s), w = bs.length * 2, h = 14;\n                cv.style.width = w + \"px\"; cv.style.height = h + \"px\"; cv.width = w * dpr; cv.height = h * dpr;\n                const g = cv.getContext(\"2d\"); g.scale(dpr, dpr);\n                g.fillStyle = \"rgba(255,255,255,0.1)\"; g.fillRect(0, 0, w, h);\n                bs.forEach((o, x) => {\n                    const t = o.c + o.w + o.r + o.u; if (!t) return;\n                    let y = 0;\n                    [[o.c, \"#22c55e\"], [o.w, \"#f43f5e\"], [o.r, \"#fbbf24\"]].forEach(([v, col]) => { if (!v) return; const hh = h * v / t; g.fillStyle = col; g.fillRect(x * 2, y, 1.6, hh); y += hh; });\n                });\n                cv._bs = bs;\n            });\n        }\n        document.addEventListener(\"mousemove\", ev => {\n            const cv = ev.target && ev.target.classList && ev.target.classList.contains(\"strip-cv\") ? ev.target : null;\n            if (!cv || !cv._bs) return;\n            const x = Math.min(cv._bs.length - 1, Math.max(0, Math.floor((ev.clientX - cv.getBoundingClientRect().left) / 2)));\n            const o = cv._bs[x]; if (!o) return;\n            const qa = DATA.questions[o.from] || {}, qb = DATA.questions[o.to] || {};\n            cv.title = \"Câu \" + (qa.id || o.from + 1) + \"–\" + (qb.id || o.to + 1) + \": \" + o.c + \" đúng · \" + o.w + \" sai · \" + o.r + \" xem đáp án · \" + o.u + \" chưa làm\";\n        });\n\n        function stripHtml(s, idx) {\n            if (total() > STRIP_BIG) { STRIP_LIST[idx] = s; return `<canvas class=\"strip-cv\" data-i=\"${idx}\" width=\"300\" height=\"14\" style=\"vertical-align:middle;\"></canvas>`; }\n            let out = \"\";\n            for (let i = 0; i < total(); i++) {\n                if (!isAsg(s, i)) continue;\n                const ch = s.states[i] || \"\";\n                const m = STRIP[ch] || [\"\", \"chưa làm\"];\n                const q = DATA.questions[i] || { id: i + 1, root: \"\" };\n                out += `<i class=\"${m[0]}\" title=\"Câu ${q.id} · ${TYPE_LABEL[qTypeOf(q)]} (${esc(qLabelOf(q))}): ${m[1]}\"></i>`;\n            }\n            return `<span class=\"strip\">${out}</span>`;\n        }\n\n        function renderStudents(list) {\n            const key = view.sortKey, dir = view.sortDir;\n            const sorted = list.slice().sort((a, b) => {\n                const va = sortValue(a, key), vb = sortValue(b, key);\n                let c;\n                if (typeof va === \"string\" || typeof vb === \"string\") c = natCmp(va, vb);\n                else c = va - vb;\n                if (c === 0) c = natCmp(a.cls, b.cls) || natCmp(a.name, b.name);\n                return c * dir;\n            });\n            const head = COLUMNS.map(c =>\n                `<th class=\"${c.key === key ? \"sorted\" : \"\"}\" ${c.nosort ? \"\" : `onclick=\"sortBy('${c.key}')\"`}>${c.label}${c.key === key ? (dir > 0 ? \" ▲\" : \" ▼\") : \"\"}</th>`\n            ).join(\"\");\n            STRIP_LIST = [];\n            const body = sorted.map((s, rowIdx) => {\n                const st = statusOf(s);\n                return `<tr>\n                    <td>${esc(s.cls)}</td>\n                    <td><b class=\"lnk\" style=\"cursor:pointer;\" data-cls=\"${esc(s.cls)}\" data-name=\"${esc(s.name)}\" onclick=\"openStudent(this.dataset.cls, this.dataset.name)\" title=\"Xem bài làm từng câu\">${esc(s.name)}</b></td>\n                    <td><span class=\"tag ${st}\">${STATUS_LABEL[st]}</span></td>\n                    <td><span class=\"bar\"><i style=\"width:${nOf(s) ? Math.round(s.done / nOf(s) * 100) : 0}%\"></i></span> ${s.done}/${nOf(s)}</td>\n                    <td class=\"num\">${s.correct}</td>\n                    <td class=\"num\">${s.wrong}</td>\n                    <td class=\"num\"><b>${s.done ? score1(s.score) : \"—\"}</b></td>\n                    <td class=\"num\">${s.attempts}</td>\n                    <td>${lvCell(s)}</td>\n                    <td>${ago(s.last)}</td>\n                    <td>${stripHtml(s, rowIdx)}</td>\n                </tr>`;\n            }).join(\"\");\n            document.getElementById(\"student-table\").innerHTML =\n                `<thead><tr>${head}</tr></thead><tbody>${body || '<tr><td colspan=\"12\" class=\"muted\" style=\"text-align:center;padding:20px;\">Không có học sinh nào khớp bộ lọc.</td></tr>'}</tbody>`;\n            if (total() > STRIP_BIG) drawStrips();\n        }\n\n        function questionStats() {\n            const scope = inClass();\n            const cl = selClasses();\n            return DATA.questions.map((q, i) => {\n                let answered = 0, wrong = 0;\n                scope.forEach(s => {\n                    const ch = s.states[i];\n                    if (isGood(ch)) answered++;\n                    else if (isBad(ch)) { answered++; wrong++; }\n                });\n                return { id: q.id, type: qTypeOf(q), root: qLabelOf(q), answer: ansFull(q), hasKey: q.hasKey, unit: q.unit || \"\", topic: q.topic || \"\", answered, wrong, rate: answered ? wrong / answered : 0, inScope: !scopeTokens(q.scope).length || cl.some(c => inScope(q.scope, c)) };\n            }).filter(x => x.inScope);\n        }\n\n        function renderTypeStats() {\n            const st = questionStats();\n            const present = TYPE_ORDER.filter(t => st.some(q => q.type === t));\n            const panel = el(\"type-panel\");\n            panel.classList.toggle(\"hidden\", present.length < 2);\n            if (present.length < 2) return;\n            el(\"type-list\").innerHTML = present.map(t => {\n                const qs = st.filter(q => q.type === t);\n                const answered = qs.reduce((a, q) => a + q.answered, 0), wrong = qs.reduce((a, q) => a + q.wrong, 0);\n                const good = answered - wrong, rate = answered ? good / answered : 0;\n                return `<div class=\"t-row\"><div><span class=\"tag t-${t}\">${TYPE_LABEL[t]}</span></div>\n                    <div class=\"muted\">${qs.length} câu · ${answered ? good + \"/\" + answered + \" lượt làm đúng\" : \"chưa có bài làm\"}</div>\n                    <div><span class=\"bar\"><i style=\"width:${Math.round(rate * 100)}%\"></i></span></div>\n                    <div class=\"num\" style=\"text-align:right;\">${answered ? pct(rate) : \"—\"}</div></div>`;\n            }).join(\"\");\n        }\n\n        const NO_UNIT = \"Chưa phân Unit\", NO_TOPIC = \"Chưa phân chủ đề\";\n        function unitRank(u) {\n            if (u === NO_UNIT) return [3, 0];\n            const m = /^unit\\s*(\\d+)/i.exec(u);\n            return m ? [0, parseInt(m[1], 10)] : [1, 0];\n        }\n        function cmpUnit(a, b) { const x = unitRank(a), y = unitRank(b); return (x[0] - y[0]) || (x[1] - y[1]) || a.localeCompare(b, \"vi\"); }\n        function setGroupBy(k) { view.groupBy = k; renderGroupStats(); }\n        function renderGroupStats() {\n            const panel = el(\"group-panel\");\n            const st = questionStats();\n            const keyOf = q => view.groupBy === \"unit\" ? (q.unit || NO_UNIT) : (q.topic || NO_TOPIC);\n            const any = st.some(q => q.unit || q.topic);\n            panel.classList.toggle(\"hidden\", !any);\n            if (!any) return;\n            el(\"grp-unit\").classList.toggle(\"on\", view.groupBy === \"unit\");\n            el(\"grp-topic\").classList.toggle(\"on\", view.groupBy === \"topic\");\n            const groups = {};\n            st.forEach(q => { (groups[keyOf(q)] = groups[keyOf(q)] || []).push(q); });\n            const none = view.groupBy === \"unit\" ? NO_UNIT : NO_TOPIC;\n            const names = Object.keys(groups).sort(view.groupBy === \"unit\" ? cmpUnit : (a, b) => (a === none) - (b === none) || a.localeCompare(b, \"vi\"));\n            el(\"group-list\").innerHTML = names.map(n => {\n                const qs = groups[n];\n                const answered = qs.reduce((a, q) => a + q.answered, 0), wrong = qs.reduce((a, q) => a + q.wrong, 0);\n                const good = answered - wrong, rate = answered ? good / answered : 0;\n                const weak = answered > 0 && rate < 0.5;\n                return `<div class=\"g-row\"><div><b>${esc(n)}</b></div>\n                    <div><span class=\"bar ${weak ? \"bad\" : \"\"}\"><i style=\"width:${Math.round(rate * 100)}%\"></i></span></div>\n                    <div class=\"muted\">${qs.length} câu · ${answered ? good + \"/\" + answered + \" lượt đúng\" : \"chưa có bài làm\"}</div>\n                    <div class=\"num\" style=\"text-align:right;\">${answered ? (weak ? \"⚠ \" : \"\") + pct(rate) : \"—\"}</div></div>`;\n            }).join(\"\");\n        }\n\n        function renderQuestions() {\n            const stats = questionStats().filter(q => q.answered > 0)\n                .sort((a, b) => b.rate - a.rate || b.wrong - a.wrong || a.id - b.id);\n            const shown = view.allQ ? stats : stats.slice(0, 10);\n            document.getElementById(\"q-list\").innerHTML = shown.length\n                ? shown.map(q => `<div class=\"q-row\">\n                        <div><b>Câu ${q.id}</b></div>\n                        <div><span class=\"tag t-${q.type}\">${TYPE_LABEL[q.type]}</span> ${esc(q.root)} <span class=\"muted\">→ ${esc(shortAns(q.answer))}</span></div>\n                        <div><span class=\"bar bad\"><i style=\"width:${Math.round(q.rate * 100)}%\"></i></span></div>\n                        <div class=\"num\" style=\"text-align:right;\">${q.wrong}/${q.answered}</div>\n                    </div>`).join(\"\")\n                : '<p class=\"muted\">Chưa có học sinh nào làm bài.</p>';\n            const t = document.getElementById(\"q-toggle\");\n            t.classList.toggle(\"hidden\", stats.length <= 10);\n            t.textContent = view.allQ ? \"Chỉ hiện 10 câu sai nhiều nhất\" : \"Xem tất cả \" + stats.length + \" câu\";\n        }\n\n        function chipsHtml(list) {\n            if (!list.length) return '<span class=\"muted\">Không có ✔️</span>';\n            return list.map(s => `<span class=\"chip\">${esc(s.name)}<small>${esc(s.cls)}</small></span>`).join(\"\");\n        }\n\n        function renderMissing() {\n            const scope = inClass();\n            const none = scope.filter(s => statusOf(s) === \"none\");\n            const idle = scope.filter(s => statusOf(s) === \"idle\");\n            document.getElementById(\"n-none\").textContent = none.length;\n            document.getElementById(\"n-idle\").textContent = idle.length;\n            document.getElementById(\"list-none\").innerHTML = chipsHtml(none);\n            document.getElementById(\"list-idle\").innerHTML = chipsHtml(idle);\n        }\n\n        // ---------------------------------------------------------------- xuất CSV\n        function buildCsv(list) {\n            const cols = [];\n            for (let i = 0; i < total(); i++) if (list.some(s => isAsg(s, i))) cols.push(i);\n            const head = [\"Lớp\", \"Họ và tên\", \"Trạng thái\", \"Đã làm\", \"Số câu giao\", \"Đúng\", \"Sai\", \"Điểm/10\", \"Tổng lượt\", \"Lần cuối\", \"Số lần rời trang\", \"Giây rời trang\", \"Số lần tự nộp\"];\n            cols.forEach(i => head.push(\"Câu \" + DATA.questions[i].id));\n            const cell = v => '\"' + String(v == null ? \"\" : v).replace(/\"/g, '\"\"') + '\"';\n            const rows = list.map(s => {\n                const r = [s.cls, s.name, STATUS_LABEL[statusOf(s)], s.done, nOf(s), s.correct, s.wrong,\n                    s.done ? String(s.score).replace(\".\", \",\") : \"\", s.attempts, s.last || \"\", s.lv || 0, s.lsec || 0, s.aus || 0];\n                cols.forEach(i => r.push(isAsg(s, i) ? (s.states[i] || \"\") : \"\"));\n                return r.map(cell).join(\",\");\n            });\n            return \"﻿\" + [head.map(cell).join(\",\")].concat(rows).join(\"\\r\\n\");\n        }\n\n        function csvSuffix() { return view.cls ? \"-\" + safeName(view.cls) : (view.khoi ? \"-khoi-\" + view.khoi : \"\"); }\n\n        function exportCsv() {\n            if (!DATA) return;\n            downloadCsv(buildCsv(filtered()), \"ket-qua-word-form-9\" + csvSuffix() + \".csv\");\n        }\n\n        // mỗi lớp đang hiển thị một file riêng (để gửi cho từng lớp / giáo viên chủ nhiệm)\n        function exportEachClass() {\n            if (!DATA) return;\n            const cls = selClasses();\n            if (!cls.length) return;\n            let i = 0;\n            const next = () => {\n                if (i >= cls.length) return;\n                const c = cls[i++];\n                downloadCsv(buildCsv(DATA.students.filter(s => s.cls === c)), \"ket-qua-word-form-9-\" + safeName(c) + \".csv\");\n                setTimeout(next, 700);\n            };\n            next();\n        }\n\n\n        // ==========================================================================\n        // PHẦN BỔ SUNG: THẺ, CÂU HỎI, LỚP HỌC, BÀI LÀM TỪNG HỌC SINH\n        // ==========================================================================\n        function enterApp() {\n            el(\"gate\").classList.add(\"hidden\");\n            el(\"app\").classList.remove(\"hidden\");\n            el(\"student-link\").href = API_URL;\n            startAutoRefresh();\n            loadQuestions();\n        }\n\n        function showTab(t) {\n            [\"results\", \"questions\", \"roster\", \"links\", \"ac\"].forEach(x => {\n                el(\"pane-\" + x).classList.toggle(\"hidden\", x !== t);\n                el(\"tab-\" + x).classList.toggle(\"active\", x === t);\n            });\n            if (t === \"questions\") loadQuestions();\n            if (t === \"roster\") loadRoster();\n            if (t === \"links\") renderLinks();\n            if (t === \"ac\") loadAc();\n        }\n\n        // gọi API có khóa giáo viên\n        async function tapi(action, extra) {\n            const r = await apiPost(Object.assign({ action: action, key: teacherKey }, extra || {}));\n            if (r && r.error === \"BAD_KEY\") { logoutTeacher(); throw new Error(\"BAD_KEY\"); }\n            return r;\n        }\n\n        function msgBox(id, text, kind) {\n            const box = el(id);\n            if (!box) return;\n            box.innerHTML = text ? `<div class=\"${kind === \"ok\" ? \"ok-box\" : kind === \"warn\" ? \"warn-box\" : \"error-box\"}\">${esc(text)}</div>` : \"\";\n        }\n\n        function openModal(html) { el(\"modal-body\").innerHTML = html; el(\"modal\").classList.remove(\"hidden\"); }\n        function closeModal() { el(\"modal\").classList.add(\"hidden\"); el(\"modal-body\").innerHTML = \"\"; modalAction = null; }\n\n        function confirmBox(text, label, fn) {\n            modalAction = fn;\n            openModal(`<h3>Xác nhận</h3><p style=\"white-space:pre-wrap;\">${esc(text)}</p>\n                <div class=\"actions\"><button class=\"btn btn-ghost\" onclick=\"closeModal()\">Hủy</button>\n                <button class=\"btn btn-danger\" onclick=\"const f = modalAction; closeModal(); if (f) f();\">${esc(label)}</button></div>`);\n        }\n\n        // ---------------------------------------------------------------- CÂU HỎI\n        async function loadQuestions() {\n            try {\n                const r = await tapi(\"tq_list\");\n                if (!r || !r.ok) { msgBox(\"tq-msg\", \"Chưa tải được danh sách câu hỏi.\"); return; }\n                TQ = r; renderQuestionsTab();\n            } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"tq-msg\", \"Không kết nối được máy chủ.\"); }\n        }\n\n        function blankHtml(sentence) { return richHtml(sentence).replace(/_{3,}/, '<span class=\"blank\">______</span>'); }\n\n        // câu q có thuộc bộ lọc khối/lớp v không (all | blank | k<khối> | c<lớp>)\n        function scopeOk(q, v) {\n            if (!v || v === \"all\") return true;\n            const toks = scopeTokens(q.scope);\n            if (v === \"blank\") return !toks.length;\n            if (!toks.length) return true;\n            if (v[0] === \"c\") return inScope(q.scope, v.slice(1));\n            const k = v.slice(1);\n            return toks.some(t => { const i = scopeInfo(t); return i.kind === \"all\" || (i.kind === \"khoi\" ? i.key === k : khoiOf(t) === k); });\n        }\n        function scopeCell(q) {\n            const toks = scopeTokens(q.scope);\n            if (!toks.length) return '<span class=\"muted\" style=\"font-size:0.78rem;\">mọi lớp</span>';\n            return toks.map(t => { const i = scopeInfo(t); return `<span class=\"scope-chip\">${i.kind === \"khoi\" ? \"Khối \" + i.key : i.kind === \"all\" ? \"Mọi lớp\" : \"Lớp \" + esc(t)}</span>`; }).join(\"\");\n        }\n        function filteredQs() {\n            const uOf = q => q.unit || NO_UNIT, tOf = q => q.topic || NO_TOPIC;\n            const needle = fold(view.qSearch.trim());\n            return (TQ ? TQ.questions : []).filter(q => (view.qType === \"all\" || (view.qType === \"bad\" ? !!q.problem : qTypeOf(q) === view.qType)) &&\n                (view.qUnit === \"all\" || uOf(q) === view.qUnit) && (view.qTopic === \"all\" || tOf(q) === view.qTopic) && scopeOk(q, view.qScope) &&\n                (!needle || fold([q.sentence, q.root, q.answer, q.alts, q.entry, q.unit, q.topic].join(\" \")).includes(needle)));\n        }\n\n        function renderQuestionsTab() {\n            if (!TQ) return;\n            const qs = TQ.questions, bad = qs.filter(q => q.problem).length;\n            el(\"tq-count\").textContent = \"(\" + qs.length + \" câu\" + (bad ? \", \" + bad + \" câu cần sửa\" : \"\") + \")\";\n            const bd = el(\"q-badge\");\n            bd.textContent = bad; bd.classList.toggle(\"hidden\", !bad);\n            el(\"seed-btn\").classList.toggle(\"hidden\", qs.length > 0);\n\n            // bộ lọc dạng bài\n            const sel = el(\"tq-type\");\n            const counts = {}; qs.forEach(q => { const t = qTypeOf(q); counts[t] = (counts[t] || 0) + 1; });\n            sel.innerHTML = '<option value=\"all\">Tất cả dạng bài (' + qs.length + ')</option>' +\n                TYPE_ORDER.map(t => `<option value=\"${t}\">${TYPE_LABEL[t]} (${counts[t] || 0})</option>`).join(\"\") +\n                (bad ? `<option value=\"bad\">⚠ Câu cần sửa (${bad})</option>` : \"\");\n            sel.value = (view.qType === \"all\" || counts[view.qType] !== undefined || (view.qType === \"bad\" && bad)) ? view.qType : \"all\";\n            if (sel.value !== view.qType) view.qType = sel.value;\n\n            // bộ lọc khối / lớp\n            const sSel = el(\"tq-scope\");\n            const kSet = {};\n            (TQ.classes || (DATA ? DATA.classes : []) || []).forEach(c => { const k = khoiOf(c); if (k) kSet[k] = 1; });\n            qs.forEach(q => scopeTokens(q.scope).forEach(t => { const i = scopeInfo(t); if (i.kind === \"khoi\") kSet[i.key] = 1; }));\n            const kList = Object.keys(kSet).sort((a, b) => a - b);\n            const cList = (TQ.classes || (DATA ? DATA.classes : []) || []);\n            const nBlank = qs.filter(q => !scopeTokens(q.scope).length).length;\n            sSel.innerHTML = '<option value=\"all\">Mọi khối / lớp (' + qs.length + ')</option>' +\n                '<option value=\"blank\">Chưa gán: mọi lớp đều thấy (' + nBlank + ')</option>' +\n                kList.map(k => `<option value=\"k${k}\">Khối ${k}: được giao (${qs.filter(q => scopeOk(q, \"k\" + k)).length})</option>`).join(\"\") +\n                cList.map(c => `<option value=\"c${esc(c)}\">Lớp ${esc(c)}: được giao (${qs.filter(q => scopeOk(q, \"c\" + c)).length})</option>`).join(\"\");\n            const okVals = [\"all\", \"blank\"].concat(kList.map(k => \"k\" + k)).concat(cList.map(c => \"c\" + c));\n            if (!okVals.includes(view.qScope)) view.qScope = \"all\";\n            sSel.value = view.qScope;\n\n            // bộ lọc Unit / chủ đề (đếm theo toàn bộ câu; chủ đề phụ thuộc Unit đang chọn)\n            const uOf = q => q.unit || NO_UNIT, tOf = q => q.topic || NO_TOPIC;\n            const uCount = {}; qs.forEach(q => { uCount[uOf(q)] = (uCount[uOf(q)] || 0) + 1; });\n            const uNames = Object.keys(uCount).sort(cmpUnit);\n            const uSel = el(\"tq-unit\");\n            uSel.innerHTML = '<option value=\"all\">Mọi Unit</option>' + uNames.map(u => `<option value=\"${esc(u)}\">${esc(u)} (${uCount[u]})</option>`).join(\"\");\n            if (view.qUnit !== \"all\" && uCount[view.qUnit] === undefined) view.qUnit = \"all\";\n            uSel.value = view.qUnit;\n            const inU = qs.filter(q => view.qUnit === \"all\" || uOf(q) === view.qUnit);\n            const tCount = {}; inU.forEach(q => { tCount[tOf(q)] = (tCount[tOf(q)] || 0) + 1; });\n            const tNames = Object.keys(tCount).sort((a, b) => (a === NO_TOPIC) - (b === NO_TOPIC) || a.localeCompare(b, \"vi\"));\n            const tSel = el(\"tq-topic\");\n            tSel.innerHTML = '<option value=\"all\">Mọi chủ đề</option>' + tNames.map(n => `<option value=\"${esc(n)}\">${esc(n)} (${tCount[n]})</option>`).join(\"\");\n            if (view.qTopic !== \"all\" && tCount[view.qTopic] === undefined) view.qTopic = \"all\";\n            tSel.value = view.qTopic;\n\n            const needle = fold(view.qSearch.trim());\n            const allFiltered = qs.filter(q => (view.qType === \"all\" || (view.qType === \"bad\" ? !!q.problem : qTypeOf(q) === view.qType)) &&\n                (view.qUnit === \"all\" || uOf(q) === view.qUnit) && (view.qTopic === \"all\" || tOf(q) === view.qTopic) && scopeOk(q, view.qScope) &&\n                (!needle || fold([q.sentence, q.root, q.answer, q.alts, q.entry, q.unit, q.topic].join(\" \")).includes(needle)));\n            const shownQs = allFiltered.slice(0, view.qLimit);\n            // chọn nhiều câu để gán Unit / chủ đề\n            Object.keys(view.sel).forEach(k => { if (!qs.some(q => String(q.id) === k)) delete view.sel[k]; });\n            const nSel = Object.keys(view.sel).length;\n            el(\"tq-sel-bar\").classList.toggle(\"hidden\", nSel === 0);\n            el(\"tq-sel-text\").innerHTML = \"Đã chọn <b>\" + nSel + \"</b> câu\";\n            const allChecked = allFiltered.length > 0 && allFiltered.every(q => view.sel[q.id]);\n\n            const head = `<thead><tr><th><input type=\"checkbox\" title=\"Chọn / bỏ chọn tất cả câu đang lọc (${allFiltered.length})\" ${allChecked ? \"checked\" : \"\"} onchange=\"toggleSelAll(this.checked)\"></th><th>#</th><th>Dạng</th><th>Câu hỏi</th><th>Từ gốc / gợi ý / mục từ</th><th>Đáp án</th><th>Unit · Chủ đề</th><th>Áp dụng cho</th><th>Trạng thái</th><th></th></tr></thead>`;\n            const body = shownQs.map(q => {\n                const t = qTypeOf(q);\n                const st = q.problem ? `<span class=\"tag bad\" title=\"${esc(q.problem)}\">⚠ ${esc(q.problem)}</span>`\n                    : q.hidden ? '<span class=\"tag off\">Đang ẩn</span>' : '<span class=\"tag ok\">Đang dùng</span>';\n                const alts = q.alts ? ` <span class=\"muted\">(hoặc ${esc(q.alts)})</span>` : \"\";\n                const opts = t === \"trac_nghiem\" && q.opts ? `<div class=\"muted\" style=\"font-size:0.8rem;margin-top:3px;\">${q.opts.map((o, i) => o ? \"ABCD\"[i] + \". \" + richHtml(o) : \"\").filter(Boolean).join(\" &nbsp;·&nbsp; \")}</div>` : \"\";\n                const fam = t === \"dien_tu\" && q.family ? `<div class=\"muted\" style=\"font-size:0.78rem;\">🌳 ${esc(q.family)}</div>` : \"\";\n                const ut = (q.unit || q.topic) ? `<div class=\"tq-unit\">${q.unit ? '<span class=\"u\">' + esc(q.unit) + '</span>' : '<span class=\"muted\">chưa gán Unit</span>'}${q.topic ? '<br>' + esc(q.topic) : \"\"}</div>` : '<span class=\"muted\" style=\"font-size:0.78rem;\">chưa gán</span>';\n                return `<tr class=\"${q.hidden ? \"hid\" : \"\"}\">\n                    <td><input type=\"checkbox\" ${view.sel[q.id] ? \"checked\" : \"\"} onchange=\"toggleSel(${q.id}, this.checked)\"></td>\n                    <td><b>${q.id}</b></td>\n                    <td>${typeTag(q)}</td>\n                    <td class=\"wrap\">${blankHtml(q.sentence)}${opts}</td>\n                    <td>${rootCell(q)}${fam}</td>\n                    <td>${q.answer ? \"<b>\" + (t === \"trac_nghiem\" ? richHtml(ansFull(Object.assign({}, q, { answerFull: \"\" }))) : esc(q.answer)) + \"</b>\" : '<span class=\"muted\">—</span>'}${alts}</td>\n                    <td>${ut}</td>\n                    <td>${scopeCell(q)}</td>\n                    <td>${st}</td>\n                    <td><div class=\"row-actions\">\n                        <button class=\"btn btn-ghost btn-sm\" title=\"Sửa câu này\" onclick=\"openQForm(${q.id})\">✏️</button>\n                        <button class=\"btn btn-ghost btn-sm\" title=\"${q.hidden ? \"Hiện lại cho học sinh\" : \"Ẩn khỏi học sinh\"}\" onclick=\"toggleHide(${q.id})\">${q.hidden ? \"👁\" : \"🙈\"}</button>\n                        <button class=\"btn btn-danger btn-sm\" title=\"Xóa câu này\" onclick=\"askDeleteQ(${q.id})\">🗑</button>\n                    </div></td></tr>`;\n            }).join(\"\");\n            const more = allFiltered.length > shownQs.length\n                ? `<tr><td colspan=\"10\" style=\"text-align:center;padding:12px;\"><button class=\"btn btn-ghost\" onclick=\"view.qLimit += 100; renderQuestionsTab();\">Hiện thêm (còn ${allFiltered.length - shownQs.length} câu)</button></td></tr>` : \"\";\n            el(\"tq-table\").innerHTML = head + \"<tbody>\" + (body || `<tr><td colspan=\"10\" class=\"muted\" style=\"text-align:center;padding:20px;\">${qs.length ? \"Không có câu nào khớp bộ lọc.\" : \"Chưa có câu hỏi. Bấm “Thêm câu”, “Dán nhiều câu” hoặc “Nạp câu mẫu”.\"}</td></tr>`) + more + \"</tbody>\";\n            el(\"tq-count\").textContent = \"(\" + qs.length + \" / \" + (TQ.maxQ || 800) + \" câu\" + (bad ? \", \" + bad + \" câu cần sửa\" : \"\") + (allFiltered.length !== qs.length ? \"; đang lọc \" + allFiltered.length : \"\") + \")\";\n\n            // mục từ điển\n            const en = TQ.entries || [];\n            const used = {}; qs.forEach(q => { if (q.entry) used[q.entry.toLowerCase()] = (used[q.entry.toLowerCase()] || 0) + 1; });\n            el(\"te-count\").textContent = \"(\" + en.length + \" mục từ)\";\n            el(\"te-table\").innerHTML = \"<thead><tr><th>Từ</th><th>Loại</th><th>Nghĩa</th><th>Ví dụ</th><th>Số câu dùng</th><th></th></tr></thead><tbody>\" +\n                (en.map(e => `<tr><td><b>${esc(e.name)}</b><div class=\"muted\" style=\"font-size:0.78rem;\">${esc(e.ipa)}</div></td><td>${esc(e.pos)}</td>\n                    <td class=\"wrap\">${esc((e.defs || []).join(\" ; \").slice(0, 110))}</td>\n                    <td class=\"wrap\">${(e.examples || []).length} câu</td>\n                    <td>${used[e.name.toLowerCase()] || 0}</td>\n                    <td><div class=\"row-actions\">\n                        <button class=\"btn btn-ghost btn-sm\" title=\"Sửa\" data-n=\"${esc(e.name)}\" onclick=\"openEForm(this.dataset.n)\">✏️</button>\n                        <button class=\"btn btn-danger btn-sm\" title=\"Xóa\" data-n=\"${esc(e.name)}\" onclick=\"askDeleteE(this.dataset.n)\">🗑</button>\n                    </div></td></tr>`).join(\"\") || '<tr><td colspan=\"6\" class=\"muted\" style=\"text-align:center;padding:16px;\">Chưa có mục từ nào. Cần có mục từ trước khi soạn câu Word entry.</td></tr>') + \"</tbody>\";\n\n            const fam = TQ.families;\n            el(\"tf-count\").textContent = \"(\" + fam.length + \" nhóm)\";\n            el(\"tf-table\").innerHTML = \"<thead><tr><th>Từ gốc</th><th>Nội dung</th><th></th></tr></thead><tbody>\" +\n                (fam.map(f => `<tr><td><b>${esc(f.name)}</b></td>\n                    <td class=\"wrap\">${esc(f.content.length > 140 ? f.content.slice(0, 140) + \"…\" : f.content)}</td>\n                    <td><div class=\"row-actions\">\n                        <button class=\"btn btn-ghost btn-sm\" title=\"Sửa\" data-n=\"${esc(f.name)}\" onclick=\"openFForm(this.dataset.n)\">✏️</button>\n                        <button class=\"btn btn-danger btn-sm\" title=\"Xóa\" data-n=\"${esc(f.name)}\" onclick=\"askDeleteF(this.dataset.n)\">🗑</button>\n                    </div></td></tr>`).join(\"\") || '<tr><td colspan=\"3\" class=\"muted\" style=\"text-align:center;padding:16px;\">Chưa có gia đình từ nào.</td></tr>') + \"</tbody>\";\n        }\n\n        const DEFAULT_TOPICS = [\"Danh từ\", \"Động từ\", \"Tính từ\", \"Trạng từ\", \"Danh động từ (V-ing)\", \"Phát âm\", \"Trọng âm\", \"Câu điều kiện\", \"Câu bị động\",\n            \"Câu tường thuật\", \"So sánh\", \"Mệnh đề quan hệ\", \"Câu ước\", \"Thì hiện tại hoàn thành\", \"Liên từ chỉ sự tương phản\", \"Cụm động từ\", \"Cụm từ cố định\", \"Giao tiếp\", \"Từ vựng\"];\n        function uniq(a) { const s = {}; return a.filter(x => x && !s[x] && (s[x] = 1)); }\n        function unitList() { return uniq((TQ.unitSuggest || []).concat(TQ.questions.map(q => q.unit))); }\n        function topicList() { return uniq(TQ.questions.map(q => q.topic).concat(DEFAULT_TOPICS)); }\n        function labelDatalists() {\n            return '<datalist id=\"dl-unit\">' + unitList().map(u => `<option value=\"${esc(u)}\">`).join(\"\") + '</datalist>' +\n                '<datalist id=\"dl-topic\">' + topicList().map(u => `<option value=\"${esc(u)}\">`).join(\"\") + '</datalist>';\n        }\n        function toggleSel(id, on) { if (on) view.sel[id] = true; else delete view.sel[id]; renderQuestionsTab(); }\n        function toggleSelAll(on) {\n            const u = el(\"tq-unit\").value, tp = el(\"tq-topic\").value, needle = fold(view.qSearch.trim());\n            TQ.questions.forEach(q => {\n                const ok = (view.qType === \"all\" || (view.qType === \"bad\" ? !!q.problem : qTypeOf(q) === view.qType)) &&\n                    (u === \"all\" || (q.unit || NO_UNIT) === u) && (tp === \"all\" || (q.topic || NO_TOPIC) === tp) && scopeOk(q, view.qScope) &&\n                    (!needle || fold([q.sentence, q.root, q.answer, q.alts, q.entry, q.unit, q.topic].join(\" \")).includes(needle));\n                if (ok) { if (on) view.sel[q.id] = true; else delete view.sel[q.id]; }\n            });\n            renderQuestionsTab();\n        }\n        function clearSel() { view.sel = {}; renderQuestionsTab(); }\n        function openTagForm() {\n            const n = Object.keys(view.sel).length;\n            if (!n) return;\n            openModal(`<h3>Gán Unit / Chủ đề cho ${n} câu đã chọn</h3>\n                <div id=\"tg-error\"></div>\n                <div class=\"form-grid\">\n                    <div class=\"full\"><label class=\"lbl\" for=\"tg-unit\">Unit</label><input type=\"text\" id=\"tg-unit\" list=\"dl-unit\" placeholder=\"Chọn trong danh sách hoặc tự gõ, VD: Unit 3: Teen Stress and Pressure\">\n                        <label class=\"muted\" style=\"display:flex;gap:6px;align-items:center;margin-top:6px;cursor:pointer;\"><input type=\"checkbox\" id=\"tg-clear-unit\"> Xóa Unit của các câu này (để trống)</label></div>\n                    <div class=\"full\"><label class=\"lbl\" for=\"tg-topic\">Chủ đề</label><input type=\"text\" id=\"tg-topic\" list=\"dl-topic\" placeholder=\"VD: Danh từ, Câu bị động, Phát âm...\">\n                        <label class=\"muted\" style=\"display:flex;gap:6px;align-items:center;margin-top:6px;cursor:pointer;\"><input type=\"checkbox\" id=\"tg-clear-topic\"> Xóa Chủ đề của các câu này (để trống)</label></div>\n                </div>\n                ${labelDatalists()}\n                <div class=\"hint\">Ô nào để trống (và không tích “Xóa”) thì giữ nguyên giá trị cũ của các câu. Gõ giống hệt nhau cho các câu cùng Unit / chủ đề để học sinh lọc được.</div>\n                <div class=\"actions\"><button class=\"btn btn-ghost\" onclick=\"closeModal()\">Hủy</button>\n                <button class=\"btn btn-accent\" id=\"tg-save\" onclick=\"applyTag()\">Gán cho ${n} câu</button></div>`);\n            setTimeout(() => { const e = el(\"tg-unit\"); if (e) e.focus(); }, 50);\n        }\n        async function applyTag() {\n            const btn = el(\"tg-save\"); btn.disabled = true;\n            try {\n                const r = await tapi(\"tq_tag\", {\n                    ids: Object.keys(view.sel).map(Number), unit: el(\"tg-unit\").value, topic: el(\"tg-topic\").value,\n                    clearUnit: el(\"tg-clear-unit\").checked, clearTopic: el(\"tg-clear-topic\").checked\n                });\n                if (r && r.ok) { TQ = r; view.sel = {}; closeModal(); renderQuestionsTab(); msgBox(\"tq-msg\", \"Đã gán Unit / Chủ đề cho \" + r.count + \" câu.\", \"ok\"); loadData(); return; }\n                el(\"tg-error\").innerHTML = `<div class=\"error-box\">${esc((r && r.message) || \"Chưa gán được. Thử lại.\")}</div>`;\n            } catch (e) { if (e.message !== \"BAD_KEY\") el(\"tg-error\").innerHTML = '<div class=\"error-box\">Không kết nối được máy chủ.</div>'; }\n            btn.disabled = false;\n        }\n\n        // ---------------------------------------------------------------- v7: gán khối / lớp cho nhiều câu\n        function classListForScope() { return (TQ && TQ.classes) || (DATA ? DATA.classes : []) || []; }\n        function openScopeForm() {\n            const n = Object.keys(view.sel).length;\n            if (!n) return;\n            const kl = allKhoi().length ? allKhoi() : [\"6\", \"7\", \"8\", \"9\"];\n            const cls = classListForScope();\n            openModal(`<h3>Gán khối / lớp cho ${n} câu đã chọn</h3>\n                <div id=\"sc-error\"></div>\n                <div class=\"full\"><label class=\"lbl\">Cách gán</label>\n                    <div class=\"chk-grid\" style=\"flex-direction:column;gap:6px;\">\n                        <label><input type=\"radio\" name=\"sc-mode\" value=\"set\" checked> <b>Chỉ dành cho</b> các khối / lớp chọn bên dưới (thay thế giá trị cũ)</label>\n                        <label><input type=\"radio\" name=\"sc-mode\" value=\"add\"> <b>Thêm</b> các khối / lớp này vào (câu dùng chung nhiều khối)</label>\n                        <label><input type=\"radio\" name=\"sc-mode\" value=\"remove\"> <b>Bớt</b> các khối / lớp này ra</label>\n                        <label><input type=\"radio\" name=\"sc-mode\" value=\"clear\"> <b>Cho mọi lớp</b> (bỏ giới hạn; mọi khối đều thấy)</label>\n                    </div></div>\n                <label class=\"lbl\">Khối</label>\n                <div class=\"chk-grid\">${kl.map(k => `<label><input type=\"checkbox\" class=\"sc-k\" value=\"${esc(k)}\"> Khối ${esc(k)}</label>`).join(\"\")}</div>\n                ${cls.length ? `<label class=\"lbl\">Lớp riêng (không bắt buộc)</label><div class=\"chk-grid\">${cls.map(c => `<label><input type=\"checkbox\" class=\"sc-c\" value=\"${esc(c)}\"> ${esc(c)}</label>`).join(\"\")}</div>` : \"\"}\n                <label class=\"lbl\" for=\"sc-extra\">Hoặc gõ thêm (cách nhau dấu | )</label>\n                <input type=\"text\" id=\"sc-extra\" placeholder=\"VD: 9/3|9/4\">\n                <div class=\"hint\">Học sinh chỉ thấy câu có khối / lớp của mình hoặc câu để “mọi lớp”. Câu dùng chung nhiều khối: chọn <b>Thêm</b> hoặc tích nhiều khối. Chọn “Bớt” cho mục cuối cùng sẽ đưa câu về “mọi lớp”.</div>\n                <div class=\"actions\"><button class=\"btn btn-ghost\" onclick=\"closeModal()\">Hủy</button>\n                <button class=\"btn btn-accent\" id=\"sc-save\" onclick=\"applyScope()\">Gán cho ${n} câu</button></div>`);\n        }\n        async function applyScope() {\n            const mode = (document.querySelector('input[name=\"sc-mode\"]:checked') || {}).value || \"set\";\n            const picks = [].slice.call(document.querySelectorAll(\".sc-k:checked,.sc-c:checked\")).map(x => x.value)\n                .concat(String(el(\"sc-extra\").value || \"\").split(\"|\").map(x => x.trim()).filter(Boolean));\n            if (mode !== \"clear\" && !picks.length) { el(\"sc-error\").innerHTML = '<div class=\"error-box\">Hãy chọn ít nhất một khối hoặc lớp.</div>'; return; }\n            const btn = el(\"sc-save\"); btn.disabled = true;\n            try {\n                const r = await tapi(\"tq_tag\", { ids: Object.keys(view.sel).map(Number), scope: picks.join(\"|\"), scopeMode: mode });\n                if (r && r.ok) {\n                    TQ = r; view.sel = {}; closeModal(); renderQuestionsTab();\n                    msgBox(\"tq-msg\", \"Đã gán khối / lớp cho \" + r.count + \" câu.\" + (r.nowAll && mode === \"remove\" ? \" (\" + r.nowAll + \" câu giờ dành cho mọi lớp.)\" : \"\") + \" Bấm “Chấm lại bài đã nộp” nếu cần cập nhật bảng điểm trong Sheet.\", \"ok\");\n                    loadData(); return;\n                }\n                el(\"sc-error\").innerHTML = `<div class=\"error-box\">${esc((r && r.message) || \"Chưa gán được. Thử lại.\")}</div>`;\n            } catch (e) { if (e.message !== \"BAD_KEY\") el(\"sc-error\").innerHTML = '<div class=\"error-box\">Không kết nối được máy chủ.</div>'; }\n            btn.disabled = false;\n        }\n\n        // xuất các câu đang lọc ra CSV theo đúng thứ tự cột của “Dán nhiều câu” / mẫu nhập (có thể gửi cho đồng nghiệp rồi dán lại)\n        function exportQuestionsCsv() {\n            if (!TQ) return;\n            const list = filteredQs();\n            if (!list.length) { msgBox(\"tq-msg\", \"Không có câu nào khớp bộ lọc để xuất.\"); return; }\n            const cell = v => '\"' + String(v == null ? \"\" : v).replace(/\"/g, '\"\"') + '\"';\n            const head = [\"Câu hỏi\", \"Từ gốc / Gợi ý\", \"Đáp án\", \"Đáp án khác\", \"Giải thích\", \"Dịch nghĩa\", \"Gia đình từ\", \"Ẩn (x)\", \"Dạng bài\", \"Lựa chọn A\", \"Lựa chọn B\", \"Lựa chọn C\", \"Lựa chọn D\", \"Mục từ điển\", \"Unit\", \"Chủ đề\", \"Áp dụng cho\"];\n            const rows = list.map(q => { const o = q.opts || [\"\", \"\", \"\", \"\"]; return [q.sentence, q.root, q.answer, q.alts, q.explain, q.translation, q.family, q.hidden ? \"x\" : \"\", TYPE_LABEL[qTypeOf(q)], o[0], o[1], o[2], o[3], q.entry, q.unit, q.topic, q.scope].map(cell).join(\",\"); });\n            const tag = view.qScope === \"all\" ? \"tat-ca\" : view.qScope === \"blank\" ? \"dung-chung\" : view.qScope[0] === \"k\" ? \"khoi-\" + view.qScope.slice(1) : \"lop-\" + safeName(view.qScope.slice(1));\n            downloadCsv(\"﻿\" + [head.map(cell).join(\",\")].concat(rows).join(\"\\r\\n\"), \"cau-hoi-word-form-\" + tag + \".csv\");\n            msgBox(\"tq-msg\", \"Đã xuất \" + list.length + \" câu.\", \"ok\");\n        }\n\n        // ---------------------------------------------------------------- v7: bảng link từng lớp / khối\n        function linkBase() { return API_URL.replace(/\\?.*$/, \"\"); }\n        function linkRows() {\n            const base = linkBase();\n            const rows = [{ name: \"Cả trường (học sinh tự chọn lớp)\", url: base, hs: DATA.students.length, n: \"—\" }];\n            allKhoi().forEach(k => {\n                const sc = DATA.students.filter(s => khoiOf(s.cls) === k);\n                const ns = {}; sc.forEach(s => { ns[nOf(s)] = 1; });\n                const nk = Object.keys(ns);\n                rows.push({ name: \"Khối \" + k, url: base + \"?khoi=\" + k, hs: sc.length, n: nk.length === 1 ? nk[0] : nk.join(\" / \") });\n            });\n            DATA.classes.forEach(c => {\n                const sc = DATA.students.filter(s => s.cls === c);\n                const shortName = c.replace(/^\\s*(lớp|lop|class)\\s+/i, \"\");   // \"LỚP 6/1\" -> \"6/1\" (máy chủ tự khớp lại)\n                rows.push({ name: \"Lớp \" + c, url: base + \"?lop=\" + encodeURIComponent(shortName).replace(/%2F/gi, \"/\"), hs: sc.length, n: sc.length ? nOf(sc[0]) : 0 });\n            });\n            return rows;\n        }\n        function renderLinks() {\n            if (!DATA) return;\n            const rows = linkRows();\n            el(\"links-table\").innerHTML = \"<thead><tr><th>Dành cho</th><th>Sĩ số</th><th>Số câu giao</th><th>Link</th><th></th></tr></thead><tbody>\" +\n                rows.map((r, i) => `<tr><td><b>${esc(r.name)}</b></td><td class=\"num\">${r.hs}</td><td class=\"num\">${esc(String(r.n))}</td>\n                    <td class=\"link-url\">${esc(r.url)}</td>\n                    <td><button class=\"btn btn-ghost btn-sm\" onclick=\"copyLink(${i})\">📋 Sao chép</button></td></tr>`).join(\"\") + \"</tbody>\";\n        }\n        async function copyText(text) {\n            try { await navigator.clipboard.writeText(text); return true; } catch (e) {}\n            try {\n                const ta = document.createElement(\"textarea\"); ta.value = text; ta.style.position = \"fixed\"; ta.style.opacity = \"0\";\n                document.body.appendChild(ta); ta.focus(); ta.select();\n                const ok = document.execCommand(\"copy\"); ta.remove(); return ok;\n            } catch (e) { return false; }\n        }\n        async function copyLink(i) {\n            const r = linkRows()[i]; if (!r) return;\n            msgBox(\"links-msg\", (await copyText(r.url)) ? \"Đã sao chép link: \" + r.name : \"Trình duyệt chưa cho sao chép. Hãy bôi đen link và sao chép thủ công.\", \"ok\");\n        }\n        async function copyAllLinks() {\n            const text = linkRows().map(r => r.name + \": \" + r.url).join(\"\\n\");\n            msgBox(\"links-msg\", (await copyText(text)) ? \"Đã sao chép tất cả link (có tên lớp) — dán vào Zalo / ghi chú.\" : \"Trình duyệt chưa cho sao chép. Hãy bôi đen bảng link và sao chép thủ công.\", \"ok\");\n        }\n\n\n        // ---------------------------------------------------------------- CHỐNG GIAN LẬN\n        let AC = null;\n        const AC_MODE_LABEL = { free: \"Tự do (không theo dõi)\", exam: \"Nhắc nhở khi Làm đề tổng hợp\", all: \"Nhắc nhở cả khi luyện tập\" };\n        async function loadAc() {\n            try {\n                const r = await tapi(\"ac_list\");\n                if (!r || !r.ok) { msgBox(\"ac-msg\", \"Chưa tải được dữ liệu chống gian lận.\"); return; }\n                AC = r; renderAc();\n            } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"ac-msg\", \"Không kết nối được máy chủ.\"); }\n        }\n        function renderAc() {\n            if (!AC) return;\n            const c = AC.cfg;\n            el(\"ac-mode\").innerHTML = [\"free\", \"exam\", \"all\"].map(m => `<option value=\"${m}\" ${c.mode === m ? \"selected\" : \"\"}>${AC_MODE_LABEL[m]}</option>`).join(\"\");\n            el(\"ac-limit\").value = c.limit;\n            const sel = (i, v) => `<select id=\"ac-cm-${i}\" style=\"background:rgba(15,23,42,0.9);color:inherit;border:1px solid var(--border);border-radius:8px;padding:6px 8px;\"><option value=\"\">Theo mặc định</option>` +\n                [\"free\", \"exam\", \"all\"].map(m => `<option value=\"${m}\" ${v === m ? \"selected\" : \"\"}>${AC_MODE_LABEL[m]}</option>`).join(\"\") + \"</select>\";\n            el(\"ac-class-table\").innerHTML = \"<thead><tr><th>Lớp</th><th>Cách làm bài</th><th>Số lần được phép</th><th>Đang áp dụng</th></tr></thead><tbody>\" +\n                (c.classes.length ? c.classes.map((x, i) => `<tr><td><b>${esc(x.name)}</b></td><td>${sel(i, x.mode)}</td>\n                    <td><input type=\"text\" id=\"ac-cl-${i}\" inputmode=\"numeric\" value=\"${x.limit || \"\"}\" placeholder=\"mặc định\" style=\"width:90px;background:rgba(15,23,42,0.9);color:inherit;border:1px solid var(--border);border-radius:8px;padding:6px 8px;\"></td>\n                    <td><span class=\"tag ${x.effMode === \"free\" ? \"off\" : \"ok\"}\">${esc(AC_MODE_LABEL[x.effMode])}${x.effMode === \"free\" ? \"\" : \" · \" + x.effLimit + \" lần\"}</span></td></tr>`).join(\"\")\n                    : '<tr><td colspan=\"4\" class=\"muted\">Chưa có lớp nào trong danh sách.</td></tr>') + \"</tbody>\";\n            const cur = el(\"ac-f-class\").value;\n            el(\"ac-f-class\").innerHTML = '<option value=\"\">Tất cả các lớp</option>' + c.classes.map(x => `<option value=\"${esc(x.name)}\" ${cur === x.name ? \"selected\" : \"\"}>${esc(x.name)}</option>`).join(\"\");\n            renderAcLog();\n        }\n        async function saveAc() {\n            if (!AC) return;\n            const lim = parseInt(el(\"ac-limit\").value, 10);\n            if (!(lim >= 1 && lim <= 20)) { msgBox(\"ac-msg\", \"Số lần được phép phải từ 1 đến 20.\"); return; }\n            const classes = AC.cfg.classes.map((x, i) => {\n                const l = parseInt(el(\"ac-cl-\" + i).value, 10);\n                return { name: x.name, mode: el(\"ac-cm-\" + i).value, limit: l >= 1 ? Math.min(20, l) : 0 };\n            });\n            try {\n                const r = await tapi(\"ac_save\", { cfg: { mode: el(\"ac-mode\").value, limit: lim, classes: classes } });\n                if (!r || !r.ok) { msgBox(\"ac-msg\", \"Chưa lưu được cài đặt. Thử lại.\"); return; }\n                AC.cfg = r.cfg; renderAc();\n                msgBox(\"ac-msg\", \"Đã lưu. Học sinh áp dụng cài đặt mới khi đăng nhập hoặc tải lại trang.\", \"ok\");\n                loadData();\n            } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"ac-msg\", \"Không kết nối được máy chủ.\"); }\n        }\n        function acFiltered() {\n            if (!AC) return [];\n            const c = el(\"ac-f-class\").value, from = el(\"ac-f-from\").value, to = el(\"ac-f-to\").value;\n            return AC.events.filter(e => (!c || e.cls === c) && (!from || e.t.slice(0, 10) >= from) && (!to || e.t.slice(0, 10) <= to));\n        }\n        function renderAcLog() {\n            if (!AC) return;\n            const ev = acFiltered();\n            const sum = {};\n            ev.forEach(e => {\n                const k = e.cls + \"|\" + e.name;\n                const o = sum[k] || (sum[k] = { cls: e.cls, name: e.name, n: 0, secs: 0, auto: 0, last: \"\" });\n                o.n++; o.secs += e.secs; if (e.auto) o.auto++; if (e.t > o.last) o.last = e.t;\n            });\n            const rows = Object.keys(sum).map(k => sum[k]).sort((a, b) => b.n - a.n || natCmp(a.name, b.name));\n            el(\"ac-sum-table\").innerHTML = \"<thead><tr><th>Lớp</th><th>Họ và tên</th><th>Số lần rời trang</th><th>Tổng thời gian rời</th><th>Bài tự nộp</th><th>Lần gần nhất</th></tr></thead><tbody>\" +\n                (rows.length ? rows.map(o => `<tr><td>${esc(o.cls)}</td><td><b>${esc(o.name)}</b></td><td class=\"num\">${o.n}</td><td>${esc(fmtSecs(o.secs))}</td>\n                    <td>${o.auto ? '<span class=\"tag bad\">' + o.auto + \"</span>\" : '<span class=\"muted\">0</span>'}</td><td>${esc(ago(o.last))}</td></tr>`).join(\"\")\n                    : '<tr><td colspan=\"6\" class=\"muted\" style=\"text-align:center;padding:16px;\">Chưa có lần rời trang nào được ghi nhận.</td></tr>') + \"</tbody>\";\n            el(\"ac-log-table\").innerHTML = \"<thead><tr><th>Thời gian</th><th>Lớp</th><th>Họ và tên</th><th>Đang ở</th><th>Lần thứ</th><th>Rời</th><th>Ghi chú</th></tr></thead><tbody>\" +\n                ev.slice(0, 200).map(e => `<tr><td>${esc(e.t.slice(0, 16).replace(\"T\", \" \"))}</td><td>${esc(e.cls)}</td><td>${esc(e.name)}</td><td>${esc(e.where)}</td><td class=\"num\">${e.n}</td>\n                    <td>${esc(fmtSecs(e.secs))}</td><td>${e.auto ? '<span class=\"tag bad\">tự nộp bài</span> ' : \"\"}${esc(e.note || \"\")}</td></tr>`).join(\"\") + \"</tbody>\";\n        }\n        function exportAcCsv() {\n            if (!AC) return;\n            const cell = v => '\"' + String(v == null ? \"\" : v).replace(/\"/g, '\"\"') + '\"';\n            const head = [\"Thời gian\", \"Lớp\", \"Họ và tên\", \"Đang ở chế độ\", \"Lần rời thứ\", \"Thời gian rời (giây)\", \"Bài bị tự nộp\", \"Ghi chú\"];\n            const rows = acFiltered().map(e => [e.t.slice(0, 19).replace(\"T\", \" \"), e.cls, e.name, e.where, e.n, e.secs, e.auto ? \"Có\" : \"\", e.note].map(cell).join(\",\"));\n            downloadCsv(\"﻿\" + [head.map(cell).join(\",\")].concat(rows).join(\"\\r\\n\"), \"nhat-ky-roi-trang.csv\");\n        }\n\n        function openQForm(id) {\n            const q = TQ && TQ.questions.find(x => x.id === id) || { id: 0, type: \"dien_tu\", sentence: \"\", root: \"\", answer: \"\", alts: \"\", explain: \"\", translation: \"\", family: \"\", hidden: false, opts: [\"\", \"\", \"\", \"\"], entry: \"\", unit: view.qUnit !== \"all\" && view.qUnit !== NO_UNIT ? view.qUnit : (view.lastUnit || \"\"), topic: view.qTopic !== \"all\" && view.qTopic !== NO_TOPIC ? view.qTopic : (view.lastTopic || \"\"), scope: view.qScope && view.qScope[0] === \"k\" ? view.qScope.slice(1) : view.qScope && view.qScope[0] === \"c\" ? view.qScope.slice(1) : (view.lastScope || \"\") };\n            const t = qTypeOf(q);\n            const famOpts = ['<option value=\"\">(không thuộc nhóm nào)</option>'].concat(\n                TQ.families.map(f => `<option value=\"${esc(f.name)}\" ${f.name.toLowerCase() === (q.family || \"\").toLowerCase() ? \"selected\" : \"\"}>${esc(f.name)}</option>`)).join(\"\");\n            const entOpts = ['<option value=\"\">(chọn mục từ điển)</option>'].concat(\n                (TQ.entries || []).map(e => `<option value=\"${esc(e.name)}\" ${e.name.toLowerCase() === (q.entry || \"\").toLowerCase() ? \"selected\" : \"\"}>${esc(e.name)}</option>`)).join(\"\");\n            const typeOpts = TYPE_ORDER.map(k => `<option value=\"${k}\" ${k === t ? \"selected\" : \"\"}>${TYPE_LABEL[k]}</option>`).join(\"\");\n            const opts = q.opts || [\"\", \"\", \"\", \"\"];\n            const letter = (q.answer || \"\").trim().toUpperCase();\n            openModal(`<h3>${id ? \"Sửa câu \" + id : \"Thêm câu hỏi\"}</h3>\n                <div id=\"qf-error\"></div>\n                <div class=\"form-grid\">\n                    <div class=\"full\"><label class=\"lbl\" for=\"qf-type\">Dạng bài</label>\n                        <select id=\"qf-type\" onchange=\"onQTypeChange()\">${typeOpts}</select>\n                        <div class=\"hint\" id=\"qf-type-hint\"></div></div>\n                    <div class=\"full\"><label class=\"lbl\" for=\"qf-sentence\" id=\"qf-sentence-lbl\">Câu hỏi</label>\n                        <textarea id=\"qf-sentence\" rows=\"3\">${esc(q.sentence)}</textarea>\n                        <div class=\"hint\" id=\"qf-sentence-hint\"></div></div>\n                    <div id=\"qf-root-wrap\"><label class=\"lbl\" for=\"qf-root\" id=\"qf-root-lbl\">Từ gốc (hiện trong ngoặc)</label><input type=\"text\" id=\"qf-root\" value=\"${esc(q.root)}\" placeholder=\"ATTRACT\"></div>\n                    <div id=\"qf-entry-wrap\"><label class=\"lbl\" for=\"qf-entry\">Mục từ điển</label><select id=\"qf-entry\">${entOpts}</select></div>\n                    <div id=\"qf-answer-wrap\"><label class=\"lbl\" for=\"qf-answer\" id=\"qf-answer-lbl\">Đáp án</label><input type=\"text\" id=\"qf-answer\" value=\"${esc(t === \"trac_nghiem\" ? letter : q.answer)}\" placeholder=\"attractions\"></div>\n                    <div class=\"full\" id=\"qf-opts-wrap\"><label class=\"lbl\">Các lựa chọn (gõ nội dung; đáp án đúng chọn ở ô bên trên)</label>\n                        <div class=\"opt-grid\">${\"ABCD\".split(\"\").map((L, i) => `<div class=\"optrow\"><b>${L}</b><input type=\"text\" id=\"qf-opt-${i}\" value=\"${esc(opts[i] || \"\")}\" placeholder=\"Lựa chọn ${L}\"></div>`).join(\"\")}</div></div>\n                    <div class=\"full\" id=\"qf-alts-wrap\"><label class=\"lbl\" for=\"qf-alts\">Đáp án khác cũng được tính đúng (không bắt buộc)</label><input type=\"text\" id=\"qf-alts\" value=\"${esc(q.alts)}\" placeholder=\"Ví dụ: prioritise  (nhiều đáp án: cách nhau dấu |)\"></div>\n                    <div class=\"full\"><label class=\"lbl\" for=\"qf-explain\">Giải thích</label><textarea id=\"qf-explain\" rows=\"3\">${esc(q.explain)}</textarea></div>\n                    <div class=\"full\"><label class=\"lbl\" for=\"qf-trans\">Dịch nghĩa</label><textarea id=\"qf-trans\" rows=\"2\">${esc(q.translation)}</textarea></div>\n                    <div id=\"qf-family-wrap\"><label class=\"lbl\" for=\"qf-family\">Gia đình từ</label><select id=\"qf-family\">${famOpts}</select></div>\n                    <div style=\"display:flex;align-items:flex-end;\"><label class=\"muted\" style=\"display:flex;gap:8px;align-items:center;cursor:pointer;\"><input type=\"checkbox\" id=\"qf-hidden\" ${q.hidden ? \"checked\" : \"\"}> Ẩn câu này khỏi học sinh</label></div>\n                    <div><label class=\"lbl\" for=\"qf-unit\">Unit (để học sinh lọc bài)</label><input type=\"text\" id=\"qf-unit\" list=\"dl-unit\" value=\"${esc(q.unit || \"\")}\" placeholder=\"VD: Unit 3: Teen Stress and Pressure\"></div>\n                    <div><label class=\"lbl\" for=\"qf-topic\">Chủ đề / chuyên đề</label><input type=\"text\" id=\"qf-topic\" list=\"dl-topic\" value=\"${esc(q.topic || \"\")}\" placeholder=\"VD: Danh từ, Câu bị động...\"></div>\n                    <div class=\"full\"><label class=\"lbl\" for=\"qf-scope\">Áp dụng cho khối / lớp (để trống = mọi lớp đều thấy)</label><input type=\"text\" id=\"qf-scope\" list=\"dl-scope\" value=\"${esc(q.scope || \"\")}\" placeholder=\"VD: 9  hoặc  6|7  hoặc  9/1|9/2\">\n                        <div class=\"hint\">Gõ số khối (6, 7, 8, 9) hoặc tên lớp như ở thẻ Lớp học. Nhiều khối / lớp cách nhau dấu | để dùng chung câu này.</div></div>\n                </div>\n                <datalist id=\"dl-scope\">${allKhoi().concat(classListForScope()).map(x => `<option value=\"${esc(x)}\">`).join(\"\")}</datalist>\n                ${labelDatalists()}\n                <div class=\"actions\"><button class=\"btn btn-ghost\" onclick=\"closeModal()\">Hủy</button>\n                <button class=\"btn btn-accent\" id=\"qf-save\" onclick=\"saveQ(${id})\">Lưu</button></div>`);\n            onQTypeChange();\n            setTimeout(() => { const e = el(\"qf-sentence\"); if (e) e.focus(); }, 50);\n        }\n\n        const TYPE_HELP = {\n            dien_tu: [\"Câu có 1 chỗ trống\", \"Gõ <b>____</b> (4 dấu gạch dưới) vào đúng chỗ học sinh phải điền. Chỉ 1 chỗ trống mỗi câu.\", \"Từ gốc (hiện trong ngoặc)\", \"ATTRACT\", \"Đáp án\", \"attractions\"],\n            trac_nghiem: [\"Câu hỏi trắc nghiệm\", \"Có thể có chỗ trống <b>____</b>. Với câu phát âm, bao phần gạch chân giữa hai dấu <b>**</b> trong lựa chọn, ví dụ ô A gõ <b>b**ea**ch</b> (hiện in đậm, gạch chân).\", \"\", \"\", \"Đáp án đúng (A, B, C hoặc D)\", \"B\"],\n            viet_lai: [\"Câu gốc cần viết lại\", \"Gõ câu gốc (không cần chỗ trống). Học sinh viết lại cả câu; so khớp không phân biệt hoa/thường và dấu câu.\", \"Gợi ý bắt đầu (không bắt buộc)\", \"Begin with: She said that ...\", \"Đáp án (cả câu đúng)\", \"She said that she was tired.\"],\n            word_entry: [\"Câu có 1 chỗ trống\", \"Gõ <b>____</b> vào chỗ học sinh phải điền cụm 2–3 từ lấy từ mục từ điển. Đáp án phải là một cụm in đậm trong mục từ điển.\", \"\", \"\", \"Đáp án (cụm 2–3 từ)\", \"put pressure on\"]\n        };\n\n        function onQTypeChange() {\n            const t = el(\"qf-type\").value, h = TYPE_HELP[t];\n            el(\"qf-sentence-lbl\").textContent = h[0];\n            el(\"qf-sentence-hint\").innerHTML = h[1];\n            el(\"qf-type-hint\").textContent = ({ dien_tu: \"Word Form: cho từ gốc, học sinh điền dạng đúng.\", trac_nghiem: \"Chọn 1 trong 4 đáp án A–D (phát âm, trọng âm, ngữ pháp, từ vựng, giao tiếp...).\", viet_lai: \"Viết lại câu sao cho nghĩa không đổi (bị động, tường thuật, điều kiện, so sánh...).\", word_entry: \"Đọc mục từ điển rồi điền cụm 2–3 từ vào câu.\" })[t];\n            el(\"qf-root-wrap\").classList.toggle(\"hidden\", !(t === \"dien_tu\" || t === \"viet_lai\"));\n            el(\"qf-root-lbl\").textContent = h[2] || \"Từ gốc\";\n            el(\"qf-root\").placeholder = h[3];\n            el(\"qf-entry-wrap\").classList.toggle(\"hidden\", t !== \"word_entry\");\n            el(\"qf-opts-wrap\").classList.toggle(\"hidden\", t !== \"trac_nghiem\");\n            el(\"qf-alts-wrap\").classList.toggle(\"hidden\", t === \"trac_nghiem\");\n            el(\"qf-family-wrap\").classList.toggle(\"hidden\", t !== \"dien_tu\");\n            el(\"qf-answer-lbl\").textContent = h[4];\n            el(\"qf-answer\").placeholder = h[5];\n            el(\"qf-answer\").maxLength = t === \"trac_nghiem\" ? 1 : 250;\n        }\n\n        function collectQ(id, hiddenOverride) {\n            const t = el(\"qf-type\").value;\n            return {\n                id: id, type: TYPE_LABEL[t], sentence: el(\"qf-sentence\").value, root: el(\"qf-root\").value, answer: el(\"qf-answer\").value,\n                alts: el(\"qf-alts\").value, explain: el(\"qf-explain\").value, translation: el(\"qf-trans\").value,\n                family: el(\"qf-family\").value, hidden: el(\"qf-hidden\").checked,\n                opts: [0, 1, 2, 3].map(i => el(\"qf-opt-\" + i).value), entry: el(\"qf-entry\").value,\n                unit: el(\"qf-unit\").value, topic: el(\"qf-topic\").value, scope: el(\"qf-scope\").value\n            };\n        }\n\n        async function saveQ(id) {\n            const btn = el(\"qf-save\"); btn.disabled = true; btn.textContent = \"Đang lưu...\";\n            try {\n                const r = await tapi(\"tq_save\", { q: collectQ(id) });\n                if (r && r.ok) { if (!id) { view.lastUnit = el(\"qf-unit\").value.trim(); view.lastTopic = el(\"qf-topic\").value.trim(); view.lastScope = el(\"qf-scope\").value.trim(); } TQ = r; closeModal(); renderQuestionsTab(); msgBox(\"tq-msg\", \"Đã lưu câu \" + r.id + \".\", \"ok\"); return; }\n                el(\"qf-error\").innerHTML = `<div class=\"error-box\">${esc((r && r.message) || \"Chưa lưu được. Thử lại.\")}</div>`;\n            } catch (e) { if (e.message !== \"BAD_KEY\") el(\"qf-error\").innerHTML = '<div class=\"error-box\">Không kết nối được máy chủ.</div>'; }\n            btn.disabled = false; btn.textContent = \"Lưu\";\n        }\n\n        async function toggleHide(id) {\n            const q = TQ.questions.find(x => x.id === id); if (!q) return;\n            try {\n                const r = await tapi(\"tq_save\", { q: Object.assign({}, q, { type: TYPE_LABEL[qTypeOf(q)], hidden: !q.hidden }) });\n                if (r && r.ok) { TQ = r; renderQuestionsTab(); msgBox(\"tq-msg\", \"\"); }\n                else msgBox(\"tq-msg\", (r && r.message) || \"Chưa đổi được.\");\n            } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"tq-msg\", \"Không kết nối được máy chủ.\"); }\n        }\n\n        function askDeleteQ(id) {\n            confirmBox(\"Xóa câu \" + id + \"? Câu này sẽ biến mất khỏi bài của học sinh và không còn tính vào điểm.\", \"Xóa câu\", async () => {\n                try {\n                    const r = await tapi(\"tq_delete\", { id: id });\n                    if (r && r.ok) { TQ = r; renderQuestionsTab(); msgBox(\"tq-msg\", \"Đã xóa câu \" + id + \".\", \"ok\"); loadData(); }\n                    else msgBox(\"tq-msg\", (r && r.message) || \"Chưa xóa được.\");\n                } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"tq-msg\", \"Không kết nối được máy chủ.\"); }\n            });\n        }\n\n        function openBulkQ() {\n            openModal(`<h3>Dán nhiều câu hỏi</h3>\n                <div id=\"bq-result\"></div>\n                <p class=\"muted\" style=\"margin-bottom:8px;\">Mỗi dòng một câu, các cột cách nhau bằng phím <b>Tab</b> (copy từ Excel / Google Sheets rồi dán vào đây). Thứ tự cột (giống sheet <b>CauHoi</b>, bỏ cột STT):</p>\n                <p style=\"margin-bottom:8px; font-size:0.85rem; line-height:1.7;\"><b>1 Câu hỏi</b> | <b>2 Từ gốc / Gợi ý</b> | <b>3 Đáp án</b> | 4 Đáp án khác | 5 Giải thích | 6 Dịch nghĩa | 7 Gia đình từ | 8 Ẩn (x) | <b>9 Dạng bài</b> | 10-13 Lựa chọn A, B, C, D | 14 Mục từ điển | <b>15 Unit</b> | <b>16 Chủ đề</b> | <b>17 Áp dụng cho (khối/lớp)</b></p>\n                <div class=\"full\" style=\"margin-bottom:8px;\"><label class=\"lbl\" for=\"bq-scope\">Áp dụng cho khối / lớp (nếu dòng không có cột 17)</label><input type=\"text\" id=\"bq-scope\" list=\"dl-scope\" value=\"${esc(view.qScope && (view.qScope[0] === \"k\" || view.qScope[0] === \"c\") ? view.qScope.slice(1) : (view.lastScope || \"\"))}\" placeholder=\"VD: 7  hoặc  6|7  (để trống = mọi lớp)\"><datalist id=\"dl-scope\">${allKhoi().concat(classListForScope()).map(x => `<option value=\"${esc(x)}\">`).join(\"\")}</datalist></div>\n                <textarea id=\"bq-text\" rows=\"11\" style=\"font-family:monospace; font-size:0.85rem;\" placeholder=\"Điền từ:&#10;She is very ____ about the exam.&#9;ANXIETY&#9;anxious&#10;&#10;Trắc nghiệm (cột 9 = Trắc nghiệm, cột 10-13 = A, B, C, D):&#10;If I ____ time, I would help.&#9;&#9;B&#9;&#9;&#9;&#9;&#9;&#9;Trắc nghiệm&#9;have&#9;had&#9;will have&#9;would have\"></textarea>\n                <div class=\"hint\">Chỉ 3 cột đầu là bắt buộc với Điền từ. Dạng bài ghi đúng chữ: Điền từ / Trắc nghiệm / Viết lại câu / Word entry (để trống = Điền từ). Trắc nghiệm: Đáp án là chữ A–D. Word entry: cột 14 phải trùng tên một mục từ đã có. Cột 15-17 (Unit, Chủ đề, Áp dụng cho) không bắt buộc: các câu cùng Unit / chủ đề nên gõ giống hệt nhau; Áp dụng cho gõ khối (6, 7, 8, 9) hoặc lớp, nhiều mục cách nhau dấu |. Câu trùng nội dung với câu đã có sẽ được bỏ qua. Tối đa 800 câu.</div>\n                <div class=\"actions\"><button class=\"btn btn-ghost\" onclick=\"closeModal()\">Đóng</button>\n                <button class=\"btn btn-accent\" id=\"bq-btn\" onclick=\"importQ()\">Thêm các câu này</button></div>`);\n        }\n\n        async function importQ() {\n            const text = el(\"bq-text\").value;\n            const btn = el(\"bq-btn\"); btn.disabled = true; btn.textContent = \"Đang thêm...\";\n            try {\n                const r = await tapi(\"tq_import\", { text: text, scope: (el(\"bq-scope\") || {}).value || \"\" });\n                if (r && r.ok) {\n                    TQ = r; renderQuestionsTab();\n                    let m = \"Đã thêm \" + r.added + \" câu.\";\n                    if (r.skipped && r.skipped.length) m += \"\\nBỏ qua \" + r.skipped.length + \" dòng:\\n\" + r.skipped.slice(0, 15).map(x => \"• Dòng \" + x.line + \": \" + x.reason).join(\"\\n\") + (r.skipped.length > 15 ? \"\\n…\" : \"\");\n                    el(\"bq-result\").innerHTML = `<div class=\"${r.skipped && r.skipped.length ? \"warn-box\" : \"ok-box\"}\">${esc(m)}</div>`;\n                    if (r.added) el(\"bq-text\").value = \"\";\n                } else el(\"bq-result\").innerHTML = `<div class=\"error-box\">${esc((r && r.message) || \"Chưa thêm được.\")}</div>`;\n            } catch (e) { if (e.message !== \"BAD_KEY\") el(\"bq-result\").innerHTML = '<div class=\"error-box\">Không kết nối được máy chủ.</div>'; }\n            btn.disabled = false; btn.textContent = \"Thêm các câu này\";\n        }\n\n        async function seedQuestions() {\n            try {\n                const r = await tapi(\"tq_seed\");\n                if (r && r.ok) { TQ = r; renderQuestionsTab(); msgBox(\"tq-msg\", \"Đã nạp bộ câu hỏi mẫu của Word Form 9.\", \"ok\"); loadData(); }\n                else msgBox(\"tq-msg\", (r && r.message) || \"Chưa nạp được.\");\n            } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"tq-msg\", \"Không kết nối được máy chủ.\"); }\n        }\n\n        async function regradeAll() {\n            const btn = el(\"regrade-btn\"); btn.disabled = true;\n            msgBox(\"tq-msg\", \"Đang chấm lại...\", \"warn\");\n            try {\n                const r = await tapi(\"regrade\");\n                if (r && r.ok) { msgBox(\"tq-msg\", \"Đã chấm lại bài của \" + r.students + \" học sinh theo đáp án hiện tại.\", \"ok\"); loadData(); }\n                else msgBox(\"tq-msg\", \"Chưa chấm lại được.\");\n            } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"tq-msg\", \"Không kết nối được máy chủ.\"); }\n            btn.disabled = false;\n        }\n\n                // ---- mục từ điển\n        function openEForm(name) {\n            const e = (TQ.entries || []).find(x => x.name === name) || { name: \"\", ipa: \"\", pos: \"\", defs: [], examples: [] };\n            openModal(`<h3>${name ? \"Sửa mục từ\" : \"Thêm mục từ điển\"}</h3>\n                <div id=\"ef-error\"></div>\n                <div class=\"form-grid\">\n                    <div><label class=\"lbl\" for=\"ef-name\">Từ cần tra</label><input type=\"text\" id=\"ef-name\" value=\"${esc(e.name)}\" placeholder=\"pressure\"></div>\n                    <div><label class=\"lbl\" for=\"ef-pos\">Từ loại</label><input type=\"text\" id=\"ef-pos\" value=\"${esc(e.pos)}\" placeholder=\"noun\"></div>\n                    <div class=\"full\"><label class=\"lbl\" for=\"ef-ipa\">Phiên âm</label><input type=\"text\" id=\"ef-ipa\" value=\"${esc(e.ipa)}\" placeholder=\"/ˈpreʃ.ər/\"></div>\n                    <div class=\"full\"><label class=\"lbl\" for=\"ef-defs\">Nghĩa tiếng Anh (mỗi nghĩa một dòng)</label><textarea id=\"ef-defs\" rows=\"3\">${esc((e.defs || []).join(\"\\n\"))}</textarea></div>\n                    <div class=\"full\"><label class=\"lbl\" for=\"ef-ex\">Câu ví dụ (mỗi câu một dòng, ít nhất 2 câu)</label><textarea id=\"ef-ex\" rows=\"6\">${esc((e.examples || []).join(\"\\n\"))}</textarea>\n                        <div class=\"hint\">Cụm 2–3 từ học sinh cần học viết giữa hai dấu <b>**</b>, ví dụ: Many students feel <b>**under pressure**</b> before exams.</div></div>\n                </div>\n                <div class=\"actions\"><button class=\"btn btn-ghost\" onclick=\"closeModal()\">Hủy</button>\n                <button class=\"btn btn-accent\" id=\"ef-save\" data-old=\"${esc(e.name)}\" onclick=\"saveE(this.dataset.old)\">Lưu</button></div>`);\n        }\n\n        async function saveE(oldName) {\n            const btn = el(\"ef-save\"); btn.disabled = true;\n            try {\n                const r = await tapi(\"te_save\", { oldName: oldName, name: el(\"ef-name\").value, ipa: el(\"ef-ipa\").value, pos: el(\"ef-pos\").value, defs: el(\"ef-defs\").value, examples: el(\"ef-ex\").value });\n                if (r && r.ok) { TQ = r; closeModal(); renderQuestionsTab(); msgBox(\"te-msg\", \"Đã lưu mục từ.\", \"ok\"); return; }\n                el(\"ef-error\").innerHTML = `<div class=\"error-box\">${esc((r && r.message) || \"Chưa lưu được.\")}</div>`;\n            } catch (e) { if (e.message !== \"BAD_KEY\") el(\"ef-error\").innerHTML = '<div class=\"error-box\">Không kết nối được máy chủ.</div>'; }\n            btn.disabled = false;\n        }\n\n        function askDeleteE(name) {\n            confirmBox(\"Xóa mục từ “\" + name + \"”? Chỉ xóa được khi không còn câu Word entry nào dùng mục từ này.\", \"Xóa mục từ\", async () => {\n                try {\n                    const r = await tapi(\"te_delete\", { name: name });\n                    if (r && r.ok) { TQ = r; renderQuestionsTab(); msgBox(\"te-msg\", \"Đã xóa mục từ.\", \"ok\"); }\n                    else msgBox(\"te-msg\", (r && r.message) || \"Chưa xóa được.\");\n                } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"te-msg\", \"Không kết nối được máy chủ.\"); }\n            });\n        }\n\n        // ---- gia đình từ\n        function openFForm(name) {\n            const f = TQ.families.find(x => x.name === name) || { name: \"\", content: \"\" };\n            openModal(`<h3>${name ? \"Sửa gia đình từ\" : \"Thêm gia đình từ\"}</h3>\n                <div id=\"ff-error\"></div>\n                <div class=\"form-grid\">\n                    <div class=\"full\"><label class=\"lbl\" for=\"ff-name\">Từ gốc (tên nhóm)</label><input type=\"text\" id=\"ff-name\" value=\"${esc(f.name)}\" placeholder=\"ATTRACT\"></div>\n                    <div class=\"full\"><label class=\"lbl\" for=\"ff-content\">Nội dung</label>\n                        <textarea id=\"ff-content\" rows=\"5\">${esc(f.content)}</textarea>\n                        <div class=\"hint\">Nối các dạng từ bằng dấu → , ví dụ: Attract (v): thu hút → attraction (n): sự thu hút → attractive (adj): hấp dẫn</div></div>\n                </div>\n                <div class=\"actions\"><button class=\"btn btn-ghost\" onclick=\"closeModal()\">Hủy</button>\n                <button class=\"btn btn-accent\" id=\"ff-save\" data-old=\"${esc(f.name)}\" onclick=\"saveF(this.dataset.old)\">Lưu</button></div>`);\n        }\n\n        async function saveF(oldName) {\n            const btn = el(\"ff-save\"); btn.disabled = true;\n            try {\n                const r = await tapi(\"tf_save\", { oldName: oldName, name: el(\"ff-name\").value, content: el(\"ff-content\").value });\n                if (r && r.ok) { TQ = r; closeModal(); renderQuestionsTab(); msgBox(\"tq-msg\", \"Đã lưu gia đình từ.\", \"ok\"); return; }\n                el(\"ff-error\").innerHTML = `<div class=\"error-box\">${esc((r && r.message) || \"Chưa lưu được.\")}</div>`;\n            } catch (e) { if (e.message !== \"BAD_KEY\") el(\"ff-error\").innerHTML = '<div class=\"error-box\">Không kết nối được máy chủ.</div>'; }\n            btn.disabled = false;\n        }\n\n        function askDeleteF(name) {\n            confirmBox(\"Xóa gia đình từ “\" + name + \"”? Các câu đang dùng nhóm này sẽ không còn gắn nhóm.\", \"Xóa nhóm\", async () => {\n                try {\n                    const r = await tapi(\"tf_delete\", { name: name });\n                    if (r && r.ok) { TQ = r; renderQuestionsTab(); msgBox(\"tq-msg\", \"Đã xóa nhóm.\", \"ok\"); }\n                    else msgBox(\"tq-msg\", (r && r.message) || \"Chưa xóa được.\");\n                } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"tq-msg\", \"Không kết nối được máy chủ.\"); }\n            });\n        }\n\n        // ---------------------------------------------------------------- LỚP HỌC\n        async function loadRoster() {\n            try {\n                const r = await tapi(\"tr_list\");\n                if (!r || !r.ok) { msgBox(\"tr-msg\", \"Chưa tải được danh sách lớp.\"); return; }\n                TR = r; renderRoster();\n            } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"tr-msg\", \"Không kết nối được máy chủ.\"); }\n        }\n\n        function renderRoster() {\n            if (!TR) return;\n            el(\"tr-count\").textContent = \"(\" + TR.students.length + \" học sinh, \" + TR.classes.length + \" lớp)\";\n            el(\"tr-cls-list\").innerHTML = TR.classes.map(c => `<option value=\"${esc(c)}\">`).join(\"\");\n            el(\"tr-list\").innerHTML = TR.classes.length ? TR.classes.map(c => {\n                const list = TR.students.filter(s => s.cls === c);\n                return `<details class=\"cls-block\" ${TR.classes.length <= 3 ? \"open\" : \"\"}><summary>Lớp ${esc(c)} — ${list.length} học sinh</summary>` +\n                    list.map((s, i) => `<div class=\"stu-row\">\n                        <span class=\"muted\">${i + 1}</span><span><b>${esc(s.name)}</b></span>\n                        <span>${s.code ? '<span class=\"tag doing\">có mã</span>' : '<span class=\"muted\">không mã</span>'}</span>\n                        <span class=\"del\">\n                            <button class=\"btn btn-ghost btn-sm\" data-c=\"${esc(s.cls)}\" data-n=\"${esc(s.name)}\" data-k=\"${esc(s.code)}\" onclick=\"openCode(this.dataset.c, this.dataset.n, this.dataset.k)\">🔑 Mã</button>\n                            <button class=\"btn btn-danger btn-sm\" data-c=\"${esc(s.cls)}\" data-n=\"${esc(s.name)}\" onclick=\"askDeleteS(this.dataset.c, this.dataset.n)\">🗑 Xóa</button>\n                        </span></div>`).join(\"\") + \"</details>\";\n            }).join(\"\") : '<p class=\"muted\">Chưa có học sinh nào. Hãy dán danh sách ở khung phía trên.</p>';\n        }\n\n        async function addRoster() {\n            const btn = el(\"tr-btn\"); btn.disabled = true; btn.textContent = \"Đang thêm...\";\n            try {\n                const r = await tapi(\"tr_import\", { cls: el(\"tr-cls\").value, text: el(\"tr-text\").value });\n                if (r && r.ok) {\n                    TR = r; renderRoster();\n                    let m = \"Đã thêm \" + r.added + \" học sinh.\";\n                    if (r.skipped && r.skipped.length) m += \"\\nĐã có sẵn, bỏ qua: \" + r.skipped.length;\n                    if (r.bad && r.bad.length) m += \"\\nKhông đọc được \" + r.bad.length + \" dòng:\\n\" + r.bad.slice(0, 8).join(\"\\n\");\n                    msgBox(\"tr-msg\", m, r.bad && r.bad.length ? \"warn\" : \"ok\");\n                    if (r.added) { el(\"tr-text\").value = \"\"; loadData(); }\n                } else msgBox(\"tr-msg\", (r && r.message) || \"Chưa thêm được.\");\n            } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"tr-msg\", \"Không kết nối được máy chủ.\"); }\n            btn.disabled = false; btn.textContent = \"Thêm vào danh sách\";\n        }\n\n        function openCode(cls, name, code) {\n            openModal(`<h3>Mã học sinh — ${esc(name)} (${esc(cls)})</h3>\n                <div id=\"cd-error\"></div>\n                <label class=\"lbl\" for=\"cd-input\">Mã (để trống = không cần mã)</label>\n                <input type=\"text\" id=\"cd-input\" value=\"${esc(code)}\" placeholder=\"Ví dụ: 4821\" style=\"width:100%;\">\n                <div class=\"actions\"><button class=\"btn btn-ghost\" onclick=\"closeModal()\">Hủy</button>\n                <button class=\"btn btn-accent\" data-c=\"${esc(cls)}\" data-n=\"${esc(name)}\" onclick=\"saveCode(this.dataset.c, this.dataset.n)\">Lưu</button></div>`);\n        }\n\n        async function saveCode(cls, name) {\n            try {\n                const r = await tapi(\"tr_code\", { cls: cls, name: name, code: el(\"cd-input\").value });\n                if (r && r.ok) { TR = r; closeModal(); renderRoster(); msgBox(\"tr-msg\", \"Đã lưu mã cho \" + name + \".\", \"ok\"); }\n                else el(\"cd-error\").innerHTML = `<div class=\"error-box\">${esc((r && r.message) || \"Chưa lưu được.\")}</div>`;\n            } catch (e) { if (e.message !== \"BAD_KEY\") el(\"cd-error\").innerHTML = '<div class=\"error-box\">Không kết nối được máy chủ.</div>'; }\n        }\n\n        function askDeleteS(cls, name) {\n            confirmBox(\"Xóa \" + name + \" (lớp \" + cls + \") khỏi danh sách? Em này sẽ không vào làm bài được nữa. Kết quả cũ vẫn còn trong Google Sheet.\", \"Xóa học sinh\", async () => {\n                try {\n                    const r = await tapi(\"tr_delete\", { cls: cls, name: name });\n                    if (r && r.ok) { TR = r; renderRoster(); msgBox(\"tr-msg\", \"Đã xóa \" + name + \".\", \"ok\"); loadData(); }\n                    else msgBox(\"tr-msg\", (r && r.message) || \"Chưa xóa được.\");\n                } catch (e) { if (e.message !== \"BAD_KEY\") msgBox(\"tr-msg\", \"Không kết nối được máy chủ.\"); }\n            });\n        }\n\n        // ---------------------------------------------------------------- BÀI LÀM TỪNG HỌC SINH + CSV TỪNG CÂU\n        async function fetchDetail() {\n            if (DETAIL) return DETAIL;\n            const r = await apiPost({ action: \"dashboard\", key: teacherKey, detail: true });\n            if (!r || !r.ok) throw new Error(\"detail\");\n            DETAIL = r;\n            return r;\n        }\n\n        const RESULT_LABEL = { \"Đ\": \"Đúng\", \"S\": \"Sai\", \"X\": \"Xem đáp án\", \"T+\": \"Đúng\", \"T-\": \"Sai\" };\n\n        async function openStudent(cls, name) {\n            openModal('<h3>Đang tải bài làm...</h3>');\n            try {\n                const d = await fetchDetail();\n                const s = d.students.find(x => x.cls === cls && x.name === name);\n                if (!s) { closeModal(); return; }\n                const rows = d.questions.map((q, i) => {\n                    if (!isAsg(s, i)) return \"\";\n                    const t = qTypeOf(q);\n                    const st = s.states[i] || \"\", txt = s.texts[i] || \"\";\n                    const good = isGood(st);\n                    let cell;\n                    if (t === \"trac_nghiem\") {\n                        cell = richHtml(q.sentence) + (st ? ` <span class=\"${good ? \"rv-ok\" : \"rv-bad\"}\">→ ${st === \"X\" ? \"(xem đáp án)\" : \"chọn \" + esc(txt || \"(bỏ trống)\")}</span>` : \"\");\n                    } else if (t === \"viet_lai\") {\n                        cell = richHtml(q.sentence) + (q.root ? ` <span class=\"muted\">(${esc(q.root)})</span>` : \"\") +\n                            (st ? `<div class=\"${good ? \"rv-ok\" : \"rv-bad\"}\">✎ ${esc(st === \"X\" ? \"(xem đáp án)\" : txt || \"(bỏ trống)\")}</div>` : \"\");\n                    } else {\n                        const filled = st ? `<span class=\"${good ? \"rv-ok\" : \"rv-bad\"}\">${esc(st === \"X\" ? q.answer : txt || \"(bỏ trống)\")}</span>` : \"______\";\n                        cell = richHtml(q.sentence).replace(/_{3,}/, filled) + ` <span class=\"muted\">(${esc(t === \"word_entry\" ? \"📖 \" + (q.entry || \"\") : q.root)})</span>`;\n                    }\n                    return `<tr><td><b>${q.id}</b></td><td>${typeTag(q)}</td><td class=\"wrap\">${cell}</td>\n                        <td>${richHtml(ansFull(q))}</td>\n                        <td>${st ? `<span class=\"${good ? \"rv-ok\" : \"rv-bad\"}\">${good ? \"✔\" : \"✘\"} ${RESULT_LABEL[st] || \"\"}</span>` : '<span class=\"muted\">chưa làm</span>'}</td></tr>`;\n                }).join(\"\");\n                openModal(`<h3>${esc(s.name)} — lớp ${esc(s.cls)}</h3>\n                    <p class=\"muted\" style=\"margin-bottom:10px;\">Đã làm ${s.done}/${nOf(s)} câu · đúng ${s.correct} · sai ${s.wrong} · điểm ${s.done ? score1(s.score) : \"—\"}/10 · lần cuối ${esc(ago(s.last))}${s.lv ? \" · rời trang \" + s.lv + \" lần (\" + esc(fmtSecs(s.lsec)) + \")\" + (s.aus ? \", \" + s.aus + \" bài tự nộp\" : \"\") : \"\"}</p>\n                    <div class=\"table-wrap\" style=\"max-height:60vh; overflow:auto;\"><table><thead><tr><th>Câu</th><th>Dạng</th><th>Bài làm của học sinh</th><th>Đáp án</th><th>Kết quả</th></tr></thead><tbody>${rows}</tbody></table></div>\n                    <div class=\"actions\"><button class=\"btn btn-ghost\" onclick=\"closeModal()\">Đóng</button></div>`);\n            } catch (e) {\n                if (e.message !== \"BAD_KEY\") openModal('<h3>Chưa tải được bài làm</h3><p>Kiểm tra mạng rồi thử lại.</p><div class=\"actions\"><button class=\"btn btn-ghost\" onclick=\"closeModal()\">Đóng</button></div>');\n            }\n        }\n\n        function downloadCsv(text, filename) {\n            const blob = new Blob([text], { type: \"text/csv;charset=utf-8\" });\n            const a = document.createElement(\"a\");\n            a.href = URL.createObjectURL(blob);\n            a.download = filename;\n            document.body.appendChild(a); a.click(); a.remove();\n            setTimeout(() => URL.revokeObjectURL(a.href), 2000);\n        }\n\n        function buildDetailCsv(d, list) {\n            const cell = v => '\"' + String(v == null ? \"\" : v).replace(/\\*\\*/g, \"\").replace(/\"/g, '\"\"') + '\"';\n            const head = [\"Lớp\", \"Họ và tên\", \"Câu\", \"Dạng bài\", \"Từ gốc / mục từ\", \"Câu hỏi\", \"Học sinh trả lời\", \"Đáp án\", \"Kết quả\", \"Điểm câu\"];\n            const key = s => s.cls + \"|\" + s.name;\n            const byKey = {}; d.students.forEach(s => { byKey[key(s)] = s; });\n            const rows = [];\n            list.forEach(f => {\n                const s = byKey[key(f)]; if (!s) return;\n                d.questions.forEach((q, i) => {\n                    if (!isAsg(s, i)) return;\n                    const st = s.states[i] || \"\";\n                    rows.push([s.cls, s.name, q.id, TYPE_LABEL[qTypeOf(q)], qTypeOf(q) === \"word_entry\" ? (q.entry || \"\") : q.root, q.sentence.replace(/_{3,}/, \"____\").replace(/\\*\\*/g, \"\"),\n                        st === \"X\" ? \"(xem đáp án)\" : (s.texts[i] || \"\"), ansFull(q),\n                        st ? (RESULT_LABEL[st] || st) : \"Chưa làm\", st ? (isGood(st) ? 1 : 0) : \"\"].map(cell).join(\",\"));\n                });\n            });\n            return \"﻿\" + [head.map(cell).join(\",\")].concat(rows).join(\"\\r\\n\");\n        }\n\n        async function exportDetailCsv() {\n            const btn = el(\"export-detail-btn\"); btn.disabled = true;\n            try {\n                const d = await fetchDetail();\n                downloadCsv(buildDetailCsv(d, filtered()), \"ket-qua-tung-cau\" + csvSuffix() + \".csv\");\n            } catch (e) { if (e.message !== \"BAD_KEY\") showLoadError(\"Chưa tải được dữ liệu từng câu. Kiểm tra mạng rồi thử lại.\"); }\n            btn.disabled = false;\n        }\n\n        // ---------------------------------------------------------------- khởi chạy\n        window.onload = async function () {\n            if (!apiConfigured()) {\n                showGateError(\"Trang này chưa được kết nối (thiếu API_URL). Hãy mở bằng link /exec?page=dashboard của Web App.\");\n                document.getElementById(\"gate-btn\").disabled = true;\n                return;\n            }\n            let saved = \"\";\n            try { saved = sessionStorage.getItem(KEY_STORE) || \"\"; } catch (e) {}\n            if (saved) {\n                teacherKey = saved;\n                const ok = await loadData(true);\n                if (ok) enterApp();\n            }\n        };\n    </script>\n</body>\n</html>\n";
var PAGE_PHATAM_ = "<!DOCTYPE html>\n<html lang=\"vi\">\n<head>\n<meta charset=\"UTF-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<title>Âm khó với người Việt - Word Form 9 by NMC</title>\n<style>\n:root{--bg:#0f172a;--card:#1e293b;--border:#334155;--text:#f1f5f9;--muted:#94a3b8;--accent:#38bdf8;--ok:#34d399;--warn:#fde68a}\n*{box-sizing:border-box;margin:0;padding:0}\nbody{background:var(--bg);color:var(--text);font-family:\"Segoe UI\",system-ui,Arial,sans-serif;line-height:1.55;padding:16px;-webkit-text-size-adjust:100%}\n.wrap{max-width:860px;margin:0 auto}\nh1{font-size:1.45rem;font-weight:800;margin-bottom:4px}\n.lead{color:var(--muted);font-size:.92rem;margin-bottom:14px}\n.bar{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:12px;margin-bottom:14px;display:flex;flex-wrap:wrap;gap:8px 14px;align-items:center}\n.bar .lb{font-weight:700;font-size:.85rem;color:var(--muted)}\n.seg{display:inline-flex;gap:6px;flex-wrap:wrap}\n.seg button{background:transparent;color:var(--text);border:1px solid var(--border);border-radius:10px;padding:6px 12px;font-size:.85rem;font-weight:700;cursor:pointer;min-height:36px}\n.seg button.on{border-color:var(--accent);color:var(--accent);background:rgba(56,189,248,.12)}\n.steps{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:8px;margin-bottom:16px}\n.step{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:10px 12px;font-size:.86rem}\n.step b{color:var(--accent)}\n.toc{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px}\n.toc a{color:var(--text);text-decoration:none;border:1px solid var(--border);border-radius:999px;padding:4px 11px;font-size:.84rem;font-weight:700}\n.card{background:var(--card);border:1px solid var(--border);border-radius:16px;padding:16px;margin-bottom:14px}\n.card h2{font-size:1.1rem;font-weight:800;display:flex;gap:10px;align-items:baseline;flex-wrap:wrap}\n.sym{font-size:1.7rem;color:var(--accent);font-weight:800}\n.tip{color:var(--muted);font-size:.9rem;margin:6px 0 10px}\n.tip b{color:var(--text)}\n.words{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px}\n.w{background:rgba(148,163,184,.1);border:1px solid var(--border);color:var(--text);border-radius:10px;padding:7px 12px;font-size:.95rem;font-weight:600;cursor:pointer;min-height:40px}\n.w:active{background:rgba(56,189,248,.25)}\n.w small{color:var(--muted);font-weight:500;margin-left:4px}\n.pairs{font-size:.86rem;color:var(--warn);margin-bottom:10px}\n.links{display:flex;flex-wrap:wrap;gap:8px}\n.links a{display:inline-block;border:1px solid var(--accent);color:var(--accent);border-radius:10px;padding:7px 12px;font-weight:700;font-size:.86rem;text-decoration:none;min-height:38px;line-height:22px}\n.foot{color:var(--muted);font-size:.8rem;margin-top:20px;border-top:1px solid var(--border);padding-top:12px}\n.foot a{color:var(--accent)}\n#msg{color:var(--warn);font-size:.84rem}#msg:empty{display:none}\n</style>\n</head>\n<body>\n<div class=\"wrap\">\n<h1>👄 Âm khó với người Việt - Lớp 7</h1>\n<p class=\"lead\">Mỗi âm có: từ mẫu để nghe (bấm vào từ), mẹo ngắn, và link xem miệng - lưỡi đặt thế nào trên Sayshape.</p>\n\n<div class=\"bar\">\n  <span class=\"lb\">Giọng:</span>\n  <span class=\"seg\" id=\"seg-acc\"><button data-k=\"acc\" data-v=\"us\">🇺🇸 Mỹ (US)</button><button data-k=\"acc\" data-v=\"uk\">🇬🇧 Anh (UK)</button></span>\n  <span class=\"lb\">Nam/Nữ:</span>\n  <span class=\"seg\" id=\"seg-sex\"><button data-k=\"sex\" data-v=\"f\">Nữ</button><button data-k=\"sex\" data-v=\"m\">Nam</button></span>\n  <span class=\"lb\">Tốc độ:</span>\n  <span class=\"seg\" id=\"seg-spd\"><button data-k=\"spd\" data-v=\"0.55\">🐢 Chậm</button><button data-k=\"spd\" data-v=\"0.75\">Vừa</button><button data-k=\"spd\" data-v=\"0.92\">Nhanh</button></span>\n  <span class=\"lb\">Lặp lại:</span>\n  <span class=\"seg\" id=\"seg-rep\"><button data-k=\"rep\" data-v=\"0\">1 lần</button><button data-k=\"rep\" data-v=\"2\">🔁 2 lần (lần 2 chậm hơn)</button></span>\n  <div id=\"msg\" style=\"flex-basis:100%\"></div>\n</div>\n\n<div class=\"steps\">\n  <div class=\"step\"><b>Bước 1.</b> Bấm từ mẫu để nghe chậm, chú ý âm màu xanh ở tiêu đề.</div>\n  <div class=\"step\"><b>Bước 2.</b> Bấm \"Xem cách đọc\" để thấy vị trí lưỡi, môi và nghe giọng thật.</div>\n  <div class=\"step\"><b>Bước 3.</b> Tự nói lại 3 lần, rồi nghe lại từ mẫu để so sánh.</div>\n</div>\n\n<div class=\"toc\" id=\"toc\"></div>\n<div id=\"list\"></div>\n\n<p class=\"foot\">Link xem cách phát âm do trang <a href=\"https://sayshape.dqnotes.com/\" target=\"_blank\" rel=\"noopener noreferrer\">Sayshape</a> (tác giả Quang, dqnotes.com) cung cấp; âm thanh và hình trên trang đó thuộc về tác giả. Trang này chỉ dẫn link, không sao chép nội dung. Phần mẹo do giáo viên tổng hợp lại bằng lời riêng. Giọng đọc mẫu ở đây là giọng máy của thiết bị.</p>\n</div>\n\n<script>\n(function(){\n  var KEY=\"wf9_nmc_voice_v1\";\n  var pref={acc:\"us\",sex:\"f\",spd:\"0.75\",rep:\"0\"};\n  try{var o=JSON.parse(localStorage.getItem(KEY)||\"null\");\n    if(o){ if(o.acc===\"uk\"||o.acc===\"us\")pref.acc=o.acc; if(o.sex===\"f\"||o.sex===\"m\")pref.sex=o.sex;\n      if([\"0.55\",\"0.75\",\"0.92\"].indexOf(String(o.spd))>=0)pref.spd=String(o.spd);\n      if(String(o.rep)===\"0\"||String(o.rep)===\"2\")pref.rep=String(o.rep);} }catch(e){}\n  function save(){ try{\n    var o={}; try{o=JSON.parse(localStorage.getItem(KEY)||\"{}\")||{};}catch(e){}\n    o.acc=pref.acc;o.sex=pref.sex;o.spd=pref.spd;o.rep=pref.rep;\n    localStorage.setItem(KEY,JSON.stringify(o)); }catch(e){} }\n\n  var SS=\"https://sayshape.dqnotes.com/\";\n  var KEYS={\"θ\":\"th\",\"ð\":\"dh\",\"ʃ\":\"sh\",\"ʒ\":\"zh\",\"s\":\"s\",\"z\":\"z\",\"l\":\"l\",\"v\":\"v\",\"d\":\"d\",\"b\":\"b\",\"g\":\"g\",\"p\":\"p\",\"t\":\"t\",\"k\":\"k\",\"iː\":\"fleece\",\"ɪ\":\"kit\"};\n  var SOUNDS=[\n   {id:\"th\",sym:\"/θ/\",title:\"th không rung\",links:[\"θ\"],\n    tip:\"Đặt đầu lưỡi chạm nhẹ sau răng cửa trên rồi thổi hơi qua khe nhỏ. KHÔNG rung cổ họng. Người Việt hay đọc thành t, s hoặc f.\",\n    words:[\"think\",\"three\",\"thank\",\"month\",\"bath\"],pairs:\"think ≠ sink · three ≠ free · thank ≠ tank\"},\n   {id:\"dh\",sym:\"/ð/\",title:\"th có rung\",links:[\"ð\"],\n    tip:\"Cùng vị trí lưỡi với /θ/ nhưng CÓ rung cổ họng (đặt tay lên cổ sẽ thấy rung). Hay bị đọc thành d hoặc z.\",\n    words:[\"this\",\"that\",\"mother\",\"brother\",\"with\"],pairs:\"they ≠ day · then ≠ den\"},\n   {id:\"sh\",sym:\"/ʃ/\",title:\"sh\",links:[\"ʃ\"],\n    tip:\"Chu môi nhẹ, lưỡi lùi hơn so với /s/, thổi hơi xì kéo dài. Một số bạn đọc thành s nên she nghe như see.\",\n    words:[\"she\",\"ship\",\"shop\",\"fish\",\"wish\"],pairs:\"she ≠ see · ship ≠ sip · sheet ≠ seat\"},\n   {id:\"zh\",sym:\"/ʒ/\",title:\"âm /ʒ/ trong measure\",links:[\"ʒ\"],\n    tip:\"Giống /ʃ/ nhưng có rung cổ họng. Hay bị đọc thành z hoặc gi.\",\n    words:[\"measure\",\"usual\",\"vision\",\"television\",\"pleasure\",\"decision\"],pairs:\"\"},\n   {id:\"fs\",sym:\"/s/ /z/\",title:\"s, z ở cuối từ\",links:[\"s\",\"z\"],\n    tip:\"Tiếng Việt không có từ kết thúc bằng âm xì nên s cuối từ rất dễ bị nuốt. Quy tắc: sau âm không rung đọc /s/ (cats, books), sau âm rung đọc /z/ (plays, friends), sau s, z, sh, ch đọc thêm âm /ɪz/ (classes, watches).\",\n    words:[\"cats\",\"books\",\"plays\",\"friends\",\"musicians\",\"classes\"],pairs:\"cat ≠ cats · book ≠ books\"},\n   {id:\"fl\",sym:\"/l/\",title:\"l ở cuối từ\",links:[\"l\"],\n    tip:\"Giữ đầu lưỡi chạm lợi sau răng trên cho đến hết âm. Đừng bỏ âm l hay đọc thành \\\"iu\\\".\",\n    words:[\"feel\",\"tell\",\"school\",\"people\",\"festival\",\"musical\"],pairs:\"feel ≠ fee · tell ≠ ten\"},\n   {id:\"fv\",sym:\"/v/\",title:\"v ở cuối từ\",links:[\"v\"],\n    tip:\"Răng trên chạm môi dưới và giữ rung cổ đến hết từ. Đừng bỏ âm v.\",\n    words:[\"five\",\"love\",\"have\",\"give\",\"live\",\"leave\"],pairs:\"live ≠ lie · leave ≠ lea\"},\n   {id:\"fd\",sym:\"/d/ /b/ /ɡ/\",title:\"d, b, g ở cuối từ\",links:[\"d\",\"b\",\"g\"],\n    tip:\"Giữ âm có rung, nguyên âm đứng trước hơi dài hơn. Đừng biến thành t, p, k.\",\n    words:[\"bed\",\"bad\",\"bag\",\"pig\",\"cab\",\"good\"],pairs:\"bed ≠ bet · bad ≠ bat · bag ≠ back · pig ≠ pick\"},\n   {id:\"pk\",sym:\"/p/ /t/ /k/\",title:\"p, t, k đầu từ\",links:[\"p\",\"t\",\"k\"],\n    tip:\"Ở đầu từ có hơi bật nhẹ (đặt tay trước miệng sẽ thấy hơi thổi ra). Thiếu hơi bật thì pen nghe như Ben, coat nghe như goat.\",\n    words:[\"pen\",\"ten\",\"coat\",\"key\",\"time\",\"park\"],pairs:\"pen ≠ Ben · ten ≠ den · coat ≠ goat\"},\n   {id:\"vow\",sym:\"/iː/ /ɪ/\",title:\"iː dài và ɪ ngắn\",links:[\"iː\",\"ɪ\"],\n    tip:\"/iː/ dài, môi dàn, lưỡi cao; /ɪ/ ngắn và thả lỏng hơn. Đừng đọc hai âm này giống nhau.\",\n    words:[\"sheep\",\"ship\",\"seat\",\"sit\",\"leave\",\"live\"],pairs:\"sheep ≠ ship · seat ≠ sit · feet ≠ fit · cheap ≠ chip\"}\n  ];\n  function esc(s){return String(s).replace(/[&<>\"']/g,function(c){return {\"&\":\"&amp;\",\"<\":\"&lt;\",\">\":\"&gt;\",'\"':\"&quot;\",\"'\":\"&#39;\"}[c];});}\n  function ssUrl(sym){ return SS+\"?sound=\"+KEYS[sym]+\"&accent=\"+(pref.acc===\"uk\"?\"uk\":\"us\")+\"&lang=vi&l1=vi\"; }\n\n  var voices=[];\n  function loadVoices(){ try{voices=window.speechSynthesis.getVoices()||[];}catch(e){voices=[];} }\n  if(\"speechSynthesis\" in window){ loadVoices(); try{window.speechSynthesis.onvoiceschanged=loadVoices;}catch(e){} }\n  var FEM=/(female|zira|hazel|susan|samantha|karen|moira|tessa|fiona|serena|kate|victoria|libby|sonia|jenny|aria|emma|amy|joanna|salli|google us english|google uk english female)/i;\n  var MAL=/(\\bmale\\b|david|mark|george|daniel|alex|fred|ryan|guy|brian|james|oliver|google uk english male)/i;\n  function gender(v){ var n=v.name||\"\"; if(/female/i.test(n))return \"f\"; if(MAL.test(n))return \"m\"; if(FEM.test(n))return \"f\"; return \"?\"; }\n  function pick(){\n    var lang=pref.acc===\"uk\"?\"en-gb\":\"en-us\";\n    var c=voices.filter(function(v){return String(v.lang||\"\").toLowerCase().replace(\"_\",\"-\")===lang;});\n    var want=c.filter(function(v){return gender(v)===pref.sex;});\n    var unk=c.filter(function(v){return gender(v)===\"?\";});\n    if(want.length)return {voice:want[0],ok:true};\n    if(unk.length)return {voice:unk[0],ok:false};\n    if(c.length)return {voice:c[0],ok:false};\n    return {voice:null,ok:false};\n  }\n  var token=0;\n  function speak(text){\n    if(!(\"speechSynthesis\" in window)){ say(\"Thiết bị này chưa hỗ trợ đọc tiếng Anh.\"); return; }\n    var tok=++token;\n    try{\n      window.speechSynthesis.cancel();\n      var r=pick(), base=parseFloat(pref.spd)||0.75;\n      function one(rate,done){\n        var u=new SpeechSynthesisUtterance(text);\n        if(r.voice){u.voice=r.voice;u.lang=r.voice.lang;}else u.lang=pref.acc===\"uk\"?\"en-GB\":\"en-US\";\n        u.rate=rate; if(done)u.onend=done; window.speechSynthesis.speak(u);\n      }\n      one(base, pref.rep===\"2\" ? function(){ if(tok!==token)return; setTimeout(function(){ if(tok===token)one(Math.max(0.45,base*0.75),null); },700); } : null);\n      say(r.voice ? \"\" : \"Thiết bị chưa có giọng \" + (pref.acc===\"uk\"?\"Anh\":\"Mỹ\") + \", đang dùng giọng mặc định.\");\n    }catch(e){}\n  }\n  function say(t){ var m=document.getElementById(\"msg\"); if(m)m.textContent=t||\"\"; }\n\n  function renderBar(){\n    [\"acc\",\"sex\",\"spd\",\"rep\"].forEach(function(k){\n      var bs=document.querySelectorAll('#seg-'+k+' button');\n      for(var i=0;i<bs.length;i++) bs[i].className = bs[i].getAttribute(\"data-v\")===pref[k] ? \"on\" : \"\";\n    });\n  }\n  function renderList(){\n    var toc=\"\", html=\"\";\n    SOUNDS.forEach(function(s){\n      toc+='<a href=\"#s-'+s.id+'\">'+esc(s.sym)+'</a>';\n      html+='<div class=\"card\" id=\"s-'+s.id+'\"><h2><span class=\"sym\">'+esc(s.sym)+'</span><span>'+esc(s.title)+'</span></h2>'+\n        '<p class=\"tip\">'+esc(s.tip)+'</p><div class=\"words\">'+\n        s.words.map(function(w){return '<button type=\"button\" class=\"w\" data-w=\"'+esc(w)+'\">🔊 '+esc(w)+'</button>';}).join(\"\")+'</div>'+\n        (s.pairs?'<p class=\"pairs\">Cặp dễ nhầm: '+esc(s.pairs)+'</p>':\"\")+\n        '<div class=\"links\">'+s.links.map(function(sym){return '<a href=\"'+esc(ssUrl(sym))+'\" target=\"_blank\" rel=\"noopener noreferrer\">▶ Xem cách đọc /'+(sym===\"g\"?\"ɡ\":esc(sym))+'/</a>';}).join(\"\")+'</div></div>';\n    });\n    document.getElementById(\"toc\").innerHTML=toc;\n    document.getElementById(\"list\").innerHTML=html;\n  }\n  document.addEventListener(\"click\",function(ev){\n    var b=ev.target.closest?ev.target.closest(\"button\"):null; if(!b)return;\n    var w=b.getAttribute(\"data-w\"); if(w){ speak(w); return; }\n    var k=b.getAttribute(\"data-k\"); if(k){ pref[k]=b.getAttribute(\"data-v\"); save(); renderBar(); renderList(); speak(k===\"rep\"||k===\"spd\"?\"slowly\":\"hello\"); }\n  });\n  renderBar(); renderList();\n})();\n</script>\n</body>\n</html>\n";


// ============================================================================
//  1. ĐIỂM VÀO WEB APP
// ============================================================================

function doGet(e) {
  _QACTIVE_ = null;   // mỗi lượt gọi đọc lại câu hỏi mới
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'classes') return json_(listClasses_());
    if (p.action === 'ping') return json_({ ok: true, service: 'Word Form 9 by NMC', time: nowIso_() });
    // Mở link /exec -> trang học sinh; /exec?page=dashboard -> bảng điều khiển giáo viên
    if (p.page === 'phatam' || p.page === 'amkho') return pagePhatam_();
    return page_(p.page === 'dashboard' || p.page === 'giaovien' ? 'dashboard' : 'student', p);
  } catch (err) {
    return json_({ ok: false, error: 'SERVER', message: String(err && err.message || err) });
  }
}

/** Địa chỉ Web App (kết thúc bằng /exec). Rỗng nếu chưa triển khai. */
function pageUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (err) { return ''; }
}

/** Trang nhỏ "Âm khó với người Việt" (không cần đăng nhập, không đọc dữ liệu học sinh). */
function pagePhatam_() {
  return HtmlService.createHtmlOutput(PAGE_PHATAM_)
    .setTitle('Âm khó với người Việt - Word Form 9 by NMC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Trả về trang học sinh hoặc trang giáo viên, tự điền sẵn địa chỉ Web App vào trang. */
function page_(which, p) {
  var isDash = which === 'dashboard';
  var html = isDash ? PAGE_DASHBOARD_ : PAGE_STUDENT_;
  var url = pageUrl_();
  if (url) {
    html = html.replace('const API_URL = "DÁN_URL_WEB_APP_VÀO_ĐÂY";', function () {
      return 'const API_URL = ' + JSON.stringify(url) + ';';
    });
  }
  if (!isDash) {
    // nhúng sẵn danh sách lớp để màn hình đăng nhập hiện ngay, không phải gọi máy chủ thêm một lượt
    try {
      var pre = JSON.stringify(presetClasses_(p)).replace(/</g, '\\u003c');
      html = html.replace('const PRELOAD_CLASSES = null;', function () { return 'const PRELOAD_CLASSES = ' + pre + ';'; });
    } catch (e) { Logger.log('PRELOAD_CLASSES: ' + e); }
  }
  return HtmlService.createHtmlOutput(html)
    .setTitle(isDash ? 'Bảng điều khiển - Word Form 9 by NMC' : 'Word Form 9 by NMC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Link riêng cho từng lớp / khối: /exec?lop=7/1 (nhiều lớp: lop=7/1,7/2) hoặc /exec?khoi=7.
 * Học sinh chỉ thấy các lớp đó ở ô chọn lớp; không khớp lớp nào thì hiện đủ các lớp như bình thường.
 */
function presetClasses_(p) {
  var all = listClasses_().classes;
  if (!p) return all;
  var lop = String(p.lop == null ? '' : p.lop).trim(), khoi = String(p.khoi == null ? '' : p.khoi).trim();
  var hit = [];
  if (lop) {
    var keys = {};
    lop.split(/[|,;]+/).forEach(function (t) { if (t.trim()) keys[classKey_(t)] = 1; });
    hit = all.filter(function (c) { return keys[classKey_(c)]; });
  } else if (khoi) {
    var ks = {};
    khoi.split(/[|,;]+/).forEach(function (t) { var d = String(t).match(/\d+/); if (d) ks[String(parseInt(d[0], 10))] = 1; });
    hit = all.filter(function (c) { return ks[khoiOf_(c)]; });
  }
  return hit.length ? hit : all;
}

function doPost(e) {
  _QACTIVE_ = null;
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'BAD_REQUEST' });
  }
  try {
    switch (req.action) {
      case 'login': return json_(login_(req));
      case 'whoami': return json_(whoami_(req));
      case 'bank': return json_(bank_(req));
      case 'submit': return json_(submit_(req));
      case 'dashboard': return json_(dashboard_(req));
      case 'resetme': return json_(resetMe_(req));
      case 'tq_list': return json_(tqList_(req));
      case 'tq_save': return json_(tqSave_(req));
      case 'tq_delete': return json_(tqDelete_(req));
      case 'tq_tag': return json_(tqTag_(req));
      case 'tq_import': return json_(tqImport_(req));
      case 'tq_seed': return json_(tqSeed_(req));
      case 'tf_save': return json_(tfSave_(req));
      case 'tf_delete': return json_(tfDelete_(req));
      case 'te_save': return json_(teSave_(req));
      case 'te_delete': return json_(teDelete_(req));
      case 'tr_list': return json_(trList_(req));
      case 'tr_import': return json_(trImport_(req));
      case 'tr_delete': return json_(trDelete_(req));
      case 'tr_code': return json_(trCode_(req));
      case 'regrade': return json_(regradeApi_(req));
      case 'ac_event': return json_(acEvent_(req));
      case 'ac_list': return json_(acList_(req));
      case 'ac_save': return json_(acSave_(req));
    }
    return json_({ ok: false, error: 'BAD_ACTION' });
  } catch (err) {
    return json_({ ok: false, error: 'SERVER', message: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
//  2. CHUẨN HÓA TÊN & ĐỐI CHIẾU DANH SÁCH LỚP
// ============================================================================

function fold_(s) {
  return String(s == null ? '' : s).normalize('NFC').replace(/\s+/g, ' ').trim().toLowerCase();
}

function strip_(s) {
  return fold_(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd');
}

function safeEq_(a, b) {
  a = String(a); b = String(b);
  var diff = a.length ^ b.length;
  var n = Math.max(a.length, b.length);
  for (var i = 0; i < n; i++) {
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }
  return diff === 0;
}

function cellText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'd/M');
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
}

function getSheet_(name) {
  return SpreadsheetApp.getActive().getSheetByName(name);
}

/** Bảo đảm sheet có ít nhất n cột (chỉ THÊM cột trống ở cuối, không đụng dữ liệu). */
function ensureCols_(sh, n) {
  if (!sh) return sh;
  var have = sh.getMaxColumns();
  if (have < n) sh.insertColumnsAfter(have, n - have);
  return sh;
}

function getProgSheet_() { return ensureCols_(getSheet_(SH_PROGRESS), PROG_COLS); }
function getQSheet_() { return ensureCols_(getSheet_(SH_QUESTIONS), Q_COLS); }

/** Đọc danh sách lớp: [{cls, name, code}] */
function readRoster_() {
  var sh = getSheet_(SH_ROSTER);
  if (!sh) return [];
  var lr = sh.getLastRow();
  if (lr < 2) return [];
  var rows = sh.getRange(2, 1, lr - 1, 3).getValues();
  var out = [];
  rows.forEach(function (r) {
    var cls = cellText_(r[0]);
    var name = cellText_(r[1]);
    if (cls && name) out.push({ cls: cls, name: name, code: cellText_(r[2]) });
  });
  return out;
}

/**
 * Tìm học sinh trong danh sách (hàm thuần, dễ kiểm thử).
 * Trả về {status: 'OK'|'NOT_FOUND'|'AMBIGUOUS'|'DUPLICATE_ROSTER', student}
 */
function matchStudent_(roster, cls, name) {
  var c = fold_(cls), n = fold_(name);
  if (!c || !n) return { status: 'NOT_FOUND' };
  var inClass = roster.filter(function (r) { return fold_(r.cls) === c; });
  var exact = inClass.filter(function (r) { return fold_(r.name) === n; });
  if (exact.length === 1) return { status: 'OK', student: exact[0] };
  if (exact.length > 1) return { status: 'DUPLICATE_ROSTER' };
  var sn = strip_(name);
  var loose = inClass.filter(function (r) { return strip_(r.name) === sn; });
  if (loose.length === 1) return { status: 'OK', student: loose[0] };
  if (loose.length > 1) return { status: 'AMBIGUOUS' };
  return { status: 'NOT_FOUND' };
}

function natCmp_(a, b) {
  var ra = String(a).match(/(\d+|\D+)/g) || [];
  var rb = String(b).match(/(\d+|\D+)/g) || [];
  var n = Math.min(ra.length, rb.length);
  for (var i = 0; i < n; i++) {
    var x = ra[i], y = rb[i];
    if (/^\d+$/.test(x) && /^\d+$/.test(y)) {
      var d = parseInt(x, 10) - parseInt(y, 10);
      if (d) return d;
    } else if (x !== y) {
      return x < y ? -1 : 1;
    }
  }
  return ra.length - rb.length;
}

function listClasses_() {
  var seen = {}, list = [];
  readRoster_().forEach(function (r) {
    if (!seen[r.cls]) { seen[r.cls] = 1; list.push(r.cls); }
  });
  list.sort(natCmp_);
  return { ok: true, classes: list };
}

// ============================================================================
//  3. MÃ XÁC THỰC (TOKEN) VÀ MẬT KHẨU GIÁO VIÊN
// ============================================================================

function getSecret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('HMAC_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('HMAC_SECRET', s);
  }
  return s;
}

function makeToken_(cls, name) {
  var sig = Utilities.computeHmacSha256Signature(fold_(cls) + '|' + fold_(name), getSecret_());
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '');
}

function authStudent_(req) {
  var m = matchStudent_(readRoster_(), req.cls, req.name);
  if (m.status !== 'OK') return null;
  var s = m.student;
  if (!safeEq_(String(req.token || ''), makeToken_(s.cls, s.name))) return null;
  return s;
}

function checkTeacherKey_(key) {
  var real = PropertiesService.getScriptProperties().getProperty('TEACHER_KEY');
  if (!real) return false;
  return safeEq_(String(key || ''), real);
}

// ============================================================================
//  4. NGÂN HÀNG CÂU HỎI (đọc từ sheet CauHoi / GiaDinhTu / TuDien)
// ============================================================================
//  Bố cục sheet CauHoi (1-based):
//   A STT | B Câu hỏi | C Từ gốc / Gợi ý | D Đáp án | E Đáp án khác (ngăn bằng |) | F Giải thích |
//   G Dịch nghĩa | H Gia đình từ | I Ẩn (gõ x) | J Dạng bài | K A | L B | M C | N D | O Mục từ điển
//  Dạng bài: Điền từ (mặc định) | Trắc nghiệm | Viết lại câu | Word entry

var TYPE_LABEL = { dien_tu: 'Điền từ', trac_nghiem: 'Trắc nghiệm', viet_lai: 'Viết lại câu', word_entry: 'Word entry' };
var TYPE_ORDER = ['dien_tu', 'trac_nghiem', 'viet_lai', 'word_entry'];
var OPT_LETTERS = ['A', 'B', 'C', 'D'];

function splitAlts_(v) {
  return String(v == null ? '' : v).split(/[|;\n]/).map(function (x) { return x.trim(); })
    .filter(function (x) { return x; });
}

/** Tách theo xuống dòng hoặc dấu | (dùng cho nghĩa và ví dụ của mục từ điển). */
function splitLines_(v) {
  return String(v == null ? '' : v).split(/\r?\n|\|/).map(function (x) { return x.trim(); })
    .filter(function (x) { return x; });
}

function str_(v) {
  return String(v == null ? '' : v).trim();
}

/** Nhận dạng "Dạng bài" từ chữ giáo viên gõ. Trả về mã dạng bài; '' nếu không nhận ra. Rỗng = Điền từ. */
function typeOf_(v) {
  var s = strip_(v);
  if (!s) return 'dien_tu';
  if (/(entry|tu dien|dictionary)/.test(s)) return 'word_entry';
  if (/(trac|mcq|multiple|choice)/.test(s)) return 'trac_nghiem';
  if (/(viet lai|rewrite|rewriting|viet cau)/.test(s)) return 'viet_lai';
  if (/(dien|word form|fill|cloze)/.test(s)) return 'dien_tu';
  return '';
}

/** Đáp án trắc nghiệm: chữ cái A-D (chấp nhận "b", "B.", "B)", "(B)"). */
function letterOf_(v) {
  var m = String(v == null ? '' : v).trim().toUpperCase().match(/^\(?([A-D])[\).:]?$/);
  return m ? m[1] : '';
}

/** Lý do một câu chưa dùng được (rỗng nếu ổn). famNames/entryNames: bảng tên hợp lệ (không bắt buộc). */
function questionProblem_(q, famNames, entryNames) {
  var blanks = (q.sentence.match(/_{3,}/g) || []).length;
  if (!q.answer) return 'Chưa có đáp án';
  if (q.type === 'trac_nghiem') {
    var filled = q.opts.filter(function (o) { return o; }).length;
    if (filled < 2) return 'Trắc nghiệm cần ít nhất 2 lựa chọn (cột K-N)';
    if (!letterOf_(q.answer)) return 'Đáp án trắc nghiệm phải là một chữ cái A, B, C hoặc D';
    if (!q.opts[OPT_LETTERS.indexOf(letterOf_(q.answer))]) return 'Đáp án trỏ tới lựa chọn đang để trống';
    return '';
  }
  if (q.type === 'viet_lai') return '';
  if (blanks === 0) return 'Câu hỏi chưa có chỗ trống (____)';
  if (blanks > 1) return 'Câu hỏi có nhiều hơn 1 chỗ trống';
  if (q.type === 'word_entry') {
    if (!q.entry) return 'Word entry chưa chọn mục từ điển (cột O)';
    if (entryNames && !entryNames[fold_(q.entry)]) return 'Mục từ điển "' + q.entry + '" chưa có trong sheet TuDien';
    return '';
  }
  if (!q.root) return 'Thiếu từ gốc';
  if (famNames && q.family && !famNames[fold_(q.family)]) return 'Gia đình từ không có trong danh sách';
  return '';
}

/** Dựng một câu hỏi từ một dòng (17 ô) của sheet CauHoi. */
function buildQuestion_(r, id) {
  var tp = typeOf_(r[9]);
  var q = {
    id: id, sentence: str_(r[1]), root: str_(r[2]), answer: str_(r[3]), alts: splitAlts_(r[4]),
    explain: str_(r[5]), translation: str_(r[6]), family: str_(r[7]), hidden: str_(r[8]) !== '',
    type: tp || 'dien_tu', typeRaw: str_(r[9]), typeBad: !tp,
    opts: [str_(r[10]), str_(r[11]), str_(r[12]), str_(r[13])], entry: str_(r[14]),
    unit: cleanLabel_(r[15]), topic: cleanLabel_(r[16]), scope: cleanScope_(r[17])
  };
  if (q.type === 'trac_nghiem' && letterOf_(q.answer)) q.answer = letterOf_(q.answer);
  return q;
}

/**
 * Đọc câu hỏi, sắp theo STT. Mặc định chỉ lấy câu đang bật VÀ dùng được.
 * includeHidden=true: lấy cả câu đang ẩn. includeInvalid=true: lấy cả câu chưa dùng được (dành cho trang giáo viên).
 */
function readQuestions_(includeHidden, includeInvalid) {
  var sh = getQSheet_();
  if (!sh) return [];
  var lr = sh.getLastRow();
  if (lr < 2) return [];
  var rows = sh.getRange(2, 1, lr - 1, Q_COLS).getValues();
  var seen = {}, out = [], entryNames = null;
  rows.forEach(function (r) {
    var id = parseInt(r[0], 10);
    var sentence = str_(r[1]);
    if (!(id >= 1 && id <= MAX_Q) || !sentence || seen[id]) return;
    var hidden = str_(r[8]) !== '';
    if (hidden && !includeHidden) return;
    seen[id] = 1;
    var q = buildQuestion_(r, id);
    if (q.type === 'word_entry' && !entryNames) {
      entryNames = {};
      readEntries_().forEach(function (e) { entryNames[fold_(e.name)] = 1; });
    }
    q.problem = q.typeBad ? 'Dạng bài "' + q.typeRaw + '" không hợp lệ (chọn: Điền từ, Trắc nghiệm, Viết lại câu, Word entry)'
      : questionProblem_(q, null, entryNames);
    if (q.problem && !includeInvalid) return;
    out.push(q);
  });
  out.sort(function (a, b) { return a.id - b.id; });
  return out;
}

/** Nhãn Unit / Chủ đề: gọn khoảng trắng, bỏ < > để hiển thị an toàn. */
function cleanLabel_(v) {
  return String(v == null ? '' : v).replace(/[<>]/g, '').replace(/\s+/g, ' ').trim().slice(0, 60);
}

/** "Áp dụng cho": chuỗi như "7", "6|7", "9/1, 9/2". Rỗng = mọi lớp đều thấy câu này. */
function cleanScope_(v) {
  return String(v == null ? '' : v).replace(/[<>]/g, '').replace(/\s+/g, ' ').trim().slice(0, 120);
}

function scopeTokens_(scope) {
  return String(scope == null ? '' : scope).split(/[|,;\n]+/).map(function (t) { return t.replace(/\s+/g, ' ').trim(); }).filter(Boolean);
}

/** Khóa so khớp tên lớp: "LỚP 7/1" -> "7/1"; "9A1" -> "9a1". */
function classKey_(s) {
  return strip_(s).replace(/^(lop|class)\s*/, '').replace(/[\s.\-_]+/g, '');
}

/** Khối của một lớp = số đầu tiên trong tên lớp ("LỚP 7/1" -> "7"; "9A2" -> "9"); rỗng nếu không có số. */
function khoiOf_(cls) {
  var m = String(cls == null ? '' : cls).match(/\d+/);
  return m ? String(parseInt(m[0], 10)) : '';
}

/** Phân loại một mục trong "Áp dụng cho": {kind:'all'} | {kind:'khoi', key:'7'} | {kind:'lop', key:'7/1'} */
function scopeTokenInfo_(t) {
  var raw = strip_(t).replace(/\s+/g, '');
  if (raw === 'tatca' || raw === 'all' || raw === 'chung' || raw === 'moilop' || raw === 'tatcacaclop' || raw === 'moikhoi') return { kind: 'all', key: 'all' };
  var m = /^(?:khoi|k)?(\d{1,2})$/.exec(raw);
  if (m) return { kind: 'khoi', key: String(parseInt(m[1], 10)) };
  return { kind: 'lop', key: classKey_(t) };
}

/** Câu có "Áp dụng cho" = scope thì học sinh lớp cls có thấy không? Trống = mọi lớp. */
function inScope_(scope, cls) {
  var toks = scopeTokens_(scope);
  if (!toks.length) return true;
  var ck = classKey_(cls), kh = khoiOf_(cls);
  for (var i = 0; i < toks.length; i++) {
    var info = scopeTokenInfo_(toks[i]);
    if (info.kind === 'all') return true;
    if (info.kind === 'khoi') { if (kh && info.key === kh) return true; continue; }
    if (info.key === ck) return true;
  }
  return false;
}

/** Trả về hàm cls -> danh sách câu lớp đó được giao (có nhớ tạm theo lớp). */
function classQuestionFilter_(qs) {
  var cache = {};
  return function (cls) {
    var k = classKey_(cls);
    if (!cache[k]) cache[k] = qs.filter(function (q) { return inScope_(q.scope, cls); });
    return cache[k];
  };
}

/** Gộp / bớt các mục "Áp dụng cho". mode: set | add | remove | clear. */
function mergeScope_(cur, add, mode) {
  if (mode === 'clear') return '';
  var seen = {}, out = [];
  var push = function (t) {
    var info = scopeTokenInfo_(t);
    if (seen[info.kind + ':' + info.key]) return;
    seen[info.kind + ':' + info.key] = 1;
    out.push(info.kind === 'khoi' ? info.key : (info.kind === 'all' ? 'Tất cả' : t));
  };
  if (mode === 'set') { scopeTokens_(add).forEach(push); return out.join('|'); }
  if (mode === 'add') { scopeTokens_(cur).forEach(push); scopeTokens_(add).forEach(push); return out.join('|'); }
  if (mode === 'remove') {
    var drop = {};
    scopeTokens_(add).forEach(function (t) { var i = scopeTokenInfo_(t); drop[i.kind + ':' + i.key] = 1; });
    scopeTokens_(cur).forEach(function (t) { var i = scopeTokenInfo_(t); if (!drop[i.kind + ':' + i.key]) push(t); });
    return out.join('|');
  }
  return cur;
}

function mapById_(qs) {
  var m = {};
  qs.forEach(function (q) { m[q.id] = q; });
  return m;
}

function readFamilies_() {
  var sh = getSheet_(SH_FAMILY);
  if (!sh) return [];
  var lr = sh.getLastRow();
  if (lr < 2) return [];
  var out = [];
  sh.getRange(2, 1, lr - 1, 2).getValues().forEach(function (r) {
    var name = str_(r[0]);
    if (name) out.push({ name: name, content: str_(r[1]) });
  });
  return out;
}

/** Mục từ điển cho dạng Word entry: [{name, ipa, pos, defs[], examples[]}]. Trong ví dụ, cụm in đậm viết **như vậy**. */
function readEntries_() {
  var sh = getSheet_(SH_ENTRY);
  if (!sh) return [];
  var lr = sh.getLastRow();
  if (lr < 2) return [];
  var out = [], seen = {};
  sh.getRange(2, 1, lr - 1, 5).getValues().forEach(function (r) {
    var name = str_(r[0]);
    if (!name || seen[fold_(name)]) return;
    seen[fold_(name)] = 1;
    out.push({ name: name, ipa: str_(r[1]), pos: str_(r[2]), defs: splitLines_(r[3]), examples: splitLines_(r[4]) });
  });
  return out;
}

function normAns_(s) {
  // Giống file gốc: bỏ khoảng trắng thừa, không phân biệt hoa/thường (không bỏ dấu câu)
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();
}

/** Chuẩn hóa "mềm" cho Viết lại câu và Word entry: bỏ dấu câu, thống nhất dấu nháy, gạch nối = khoảng trắng. */
function normLoose_(s) {
  return String(s == null ? '' : s).normalize('NFC').toLowerCase()
    .replace(/[’‘`´]/g, "'")
    .replace(/[^a-z0-9À-ɏ\s']/g, ' ')
    .replace(/\s+/g, ' ').trim();
}

function gradeAnswer_(q, text) {
  if (!str_(text) || !q.answer) return false;
  if (q.type === 'trac_nghiem') return letterOf_(text) === letterOf_(q.answer);
  var norm = (q.type === 'viet_lai' || q.type === 'word_entry') ? normLoose_ : normAns_;
  var t = norm(text);
  if (!t) return false;
  if (t === norm(q.answer)) return true;
  return q.alts.some(function (a) { return norm(a) === t; });
}

/** Trạng thái cho một bài làm: Đ / S, hoặc ? nếu câu chưa có đáp án. */
function statusFor_(q, text) {
  if (!str_(text)) return ST_BAD;
  if (!q.answer) return ST_PEND;
  return gradeAnswer_(q, text) ? ST_OK : ST_BAD;
}

function isGood_(st) { return st === ST_OK || st === ST_SELF_OK; }

/** Đáp án hiển thị cho học sinh. Trắc nghiệm: "B. nội dung lựa chọn". */
function answerText_(q) {
  if (q.type === 'trac_nghiem') {
    var i = OPT_LETTERS.indexOf(letterOf_(q.answer));
    return i >= 0 ? q.answer + '. ' + q.opts[i] : q.answer;
  }
  return q.answer + (q.alts.length ? ' / ' + q.alts[0] : '');
}

/** Nhãn ngắn của câu (dùng cho nhật ký, thống kê). */
function qLabel_(q) {
  if (q.type === 'dien_tu') return q.root;
  if (q.type === 'word_entry') return 'Word entry: ' + q.entry;
  return TYPE_LABEL[q.type] + ': ' + q.sentence.slice(0, 40);
}

/** Thông tin trả cho học sinh sau khi làm. Đáp án chỉ lộ khi đã làm/xem. */
function infoFor_(q, st) {
  var showAns = q.answer && (st === ST_OK || st === ST_BAD || st === ST_REV);
  return {
    hasKey: !!q.answer,
    answer: showAns ? answerText_(q) : '',
    answerMain: showAns ? q.answer : '',
    explain: q.explain,
    translation: q.translation
  };
}

/**
 * Ngân hàng câu hỏi cho học sinh, đúng dạng dữ liệu của file gốc (questionsData + wordFamilies),
 * kèm đáp án để trang chấm ngay trên trình duyệt như file gốc.
 */
function bank_(req) {
  var s = authStudent_(req);
  if (!s) return { ok: false, error: 'AUTH' };
  var b = bankData_(s.cls);
  b.ok = true;
  return b;
}

/** Dữ liệu câu hỏi cho học sinh (không kiểm tra đăng nhập: chỉ gọi sau khi đã xác thực). */
function bankData_(cls) {
  var fams = readFamilies_();
  var famId = {};
  fams.forEach(function (f, i) { if (!famId[fold_(f.name)]) famId[fold_(f.name)] = i + 1; });
  var entries = readEntries_();
  var entId = {};
  entries.forEach(function (e, i) { entId[fold_(e.name)] = i + 1; });
  var questions = readQuestionsActive_().filter(function (q) { return !cls || inScope_(q.scope, cls); }).map(function (q) {
    var o = {
      id: q.id, type: q.type, sentence: q.sentence.replace(/_{3,}/g, '__________'), root: q.root, answer: q.answer,
      explain: q.explain, translation: q.translation, familyId: famId[fold_(q.family)] || 0
    };
    if (q.unit) o.unit = q.unit;
    if (q.topic) o.topic = q.topic;
    if (q.alts.length) o.altAnswers = q.alts;
    if (q.type === 'trac_nghiem') o.opts = q.opts;
    if (q.type === 'word_entry') o.entryId = entId[fold_(q.entry)] || 0;
    return o;
  });
  return {
    questions: questions,
    families: fams.map(function (f, i) { return { id: i + 1, root: f.name, content: f.content }; }),
    entries: entries.map(function (e, i) { return { id: i + 1, name: e.name, ipa: e.ipa, pos: e.pos, defs: e.defs, examples: e.examples }; })
  };
}

// ============================================================================
//  5. ĐĂNG NHẬP / KHÔI PHỤC PHIÊN
// ============================================================================

function login_(req) {
  var m = matchStudent_(readRoster_(), req.cls, req.name);
  if (m.status !== 'OK') return { ok: false, error: m.status };
  var s = m.student;
  if (s.code) {
    var given = String(req.code || '').trim();
    if (!given) return { ok: false, error: 'NEED_CODE' };
    if (given.toLowerCase() !== s.code.toLowerCase()) {
      Utilities.sleep(600);
      return { ok: false, error: 'BAD_CODE' };
    }
  }
  withLock_(function () { ensureProgressRow_(s.cls, s.name); });
  return {
    ok: true, cls: s.cls, name: s.name,
    token: makeToken_(s.cls, s.name),
    progress: getProgress_(s.cls, s.name),
    bank: bankData_(s.cls),    // gửi kèm để học sinh vào bài ngay, đỡ một lượt gọi máy chủ
    ac: acFor_(s.cls)
  };
}

function whoami_(req) {
  var s = authStudent_(req);
  if (!s) return { ok: false, error: 'AUTH' };
  return { ok: true, cls: s.cls, name: s.name, progress: getProgress_(s.cls, s.name), bank: bankData_(s.cls), ac: acFor_(s.cls) };
}

// ============================================================================
//  6. GHI KẾT QUẢ, CHẤM ĐIỂM
// ============================================================================

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function findProgressRow_(sh, cls, name) {
  var lr = sh.getLastRow();
  if (lr < 2) return 0;
  var c = fold_(cls), n = fold_(name);
  var rows = sh.getRange(2, 1, lr - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (fold_(cellText_(rows[i][0])) === c && fold_(rows[i][1]) === n) return i + 2;
  }
  return 0;
}

function ensureProgressRow_(cls, name) {
  var sh = getProgSheet_();
  if (!sh) throw new Error('Chưa thiết lập: hãy chạy menu Word Form 9 > 1. Thiết lập ban đầu');
  var row = findProgressRow_(sh, cls, name);
  if (row) return row;
  var blank = [cls, name, 0, 0, 0, 0, 0, 0, 0, 0, ''];
  var W = progWidth_();
  while (blank.length < W) blank.push('');
  blank[COL_TEXTS - 1] = '{}';
  row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, W).setValues([blank]);
  return row;
}

/** Tính lại các cột tổng hợp C..I của một dòng TienDo. */
function summarizeRow_(vals, activeIds) {
  var done = 0, good = 0, bad = 0, pend = 0;
  activeIds.forEach(function (id) {
    var v = vals[qCol_(id) - 1];
    if (!v) return;
    done++;
    if (isGood_(v)) good++;
    else if (v === ST_PEND) pend++;
    else bad++;
  });
  var n = activeIds.length;
  vals[2] = done;
  vals[3] = good;
  vals[4] = bad;
  vals[5] = pend;
  vals[6] = n ? done / n : 0;
  vals[7] = (good + bad) ? good / (good + bad) : 0;
  vals[8] = n ? Math.round(good / n * 100) / 10 : 0;
}

function parseTexts_(cell) {
  try { return JSON.parse(cell || '{}') || {}; } catch (e) { return {}; }
}

/**
 * Ô JSON bài làm của TienDo là MỘT ô Google Sheets (giới hạn 50.000 ký tự). Với hàng nghìn câu,
 * nếu vượt ngưỡng an toàn thì cắt ngắn dần phần chữ học sinh gõ (Đ/S/X ở từng cột câu vẫn giữ nguyên)
 * và ghi lại STT bị cắt trong "_cut" để chấm lại không chấm sai các câu đó.
 */
var TEXTS_MAX = 45000;
function packTexts_(texts) {
  var s = JSON.stringify(texts);
  if (s.length <= TEXTS_MAX) return s;
  var cut = (texts._cut && texts._cut.slice) ? texts._cut.slice() : [];
  var limits = [120, 60, 30, 12, 0];
  for (var li = 0; li < limits.length; li++) {
    var L = limits[li], copy = {}, marked = {};
    cut.forEach(function (id) { marked[id] = 1; });
    Object.keys(texts).forEach(function (k) {
      if (k === '_cut') return;
      var t = str_(texts[k]);
      if (t.length > L) { t = t.slice(0, L); if (!marked[k]) { marked[k] = 1; } }
      copy[k] = t;
    });
    copy._cut = Object.keys(marked).map(function (k) { return parseInt(k, 10); });
    s = JSON.stringify(copy);
    if (s.length <= TEXTS_MAX) return s;
  }
  return '{}';
}

/** Số cột TienDo thật sự cần đọc/ghi: đủ tới STT lớn nhất đang có trong CauHoi (tối thiểu hết cột JSON). */
function progWidth_() {
  var sh = getQSheet_();
  var lr = sh ? sh.getLastRow() : 0;
  var maxId = 1;
  if (lr >= 2) {
    sh.getRange(2, 1, lr - 1, 1).getValues().forEach(function (r) {
      var id = parseInt(r[0], 10);
      if (id > maxId && id <= MAX_Q) maxId = id;
    });
  }
  return Math.max(COL_TEXTS, qCol_(maxId));
}

/** Danh sách câu đang bật, đọc 1 lần cho mỗi lượt gọi (chỉ dùng ở các hàm chỉ-đọc/học sinh). */
var _QACTIVE_ = null;
function readQuestionsActive_() {
  if (!_QACTIVE_) _QACTIVE_ = readQuestions_(false);
  return _QACTIVE_;
}

function getProgress_(cls, name) {
  var sh = getProgSheet_();
  var row = sh ? findProgressRow_(sh, cls, name) : 0;
  var out = { states: {}, texts: {}, info: {} };
  if (!row) return out;
  var vals = sh.getRange(row, 1, 1, progWidth_()).getValues()[0];
  var texts = parseTexts_(vals[COL_TEXTS - 1]);
  readQuestionsActive_().forEach(function (q) {
    if (!inScope_(q.scope, cls)) return;
    var st = vals[qCol_(q.id) - 1];
    if (!st) return;
    out.states[q.id] = st;
    out.texts[q.id] = texts[q.id] || '';
    out.info[q.id] = infoFor_(q, st);
  });
  return out;
}

function modeLabel_(mode, item) {
  if (item.revealed) return 'Xem đáp án';
  if (mode === 'worksheet') return 'Làm đề tổng hợp';
  return 'Luyện từng câu';
}

function submit_(req) {
  var s = authStudent_(req);
  if (!s) return { ok: false, error: 'AUTH' };
  var items = req.items;
  if (!Array.isArray(items) || !items.length || items.length > MAX_Q) {
    return { ok: false, error: 'BAD_ITEMS' };
  }
  var batchId = String(req.batchId || '').slice(0, 60);
  var mode = String(req.mode || '');
  var questions = readQuestionsActive_().filter(function (q) { return inScope_(q.scope, s.cls); });   // chỉ câu lớp này được giao
  var qmap = mapById_(questions);
  var activeIds = questions.map(function (q) { return q.id; });

  return withLock_(function () {
    var cache = CacheService.getScriptCache();
    var dup = !!(batchId && cache.get('b_' + batchId));
    var ss = SpreadsheetApp.getActive();
    var sh = getProgSheet_();
    var row = ensureProgressRow_(s.cls, s.name);
    var W = progWidth_();
    var vals = sh.getRange(row, 1, 1, W).getValues()[0];
    var texts = parseTexts_(vals[COL_TEXTS - 1]);

    var now = new Date();
    var logRows = [], results = [];
    items.forEach(function (it) {
      var qid = parseInt(it && it.qid, 10);
      var q = qmap[qid];
      if (!q || qCol_(qid) > W) return;
      var text = str_(it.text).slice(0, 250);
      var st;
      if (it.revealed) {
        if (!q.answer) { results.push({ qid: qid, status: null, noKey: true }); return; }
        st = ST_REV; text = q.answer;
      } else {
        st = statusFor_(q, text);
      }
      var info = infoFor_(q, st);
      info.qid = qid; info.status = st; info.text = text;
      results.push(info);
      if (dup) return;
      vals[qCol_(qid) - 1] = st;
      texts[qid] = text;
      if (texts._cut) texts._cut = texts._cut.filter(function (x) { return x !== qid; });
      logRows.push([now, s.cls, s.name, modeLabel_(mode, it), qid, qLabel_(q), text, st, batchId]);
    });

    if (!dup && logRows.length) {
      summarizeRow_(vals, activeIds);
      vals[9] = (Number(vals[9]) || 0) + logRows.length;
      vals[10] = now;
      vals[COL_TEXTS - 1] = packTexts_(texts);
      sh.getRange(row, 1, 1, W).setValues([vals]);
      var log = ss.getSheetByName(SH_LOG);
      if (log) log.getRange(log.getLastRow() + 1, 1, logRows.length, 9).setValues(logRows);
      if (batchId) cache.put('b_' + batchId, '1', 21600);
    }
    return { ok: true, results: results };
  });
}

/** Học sinh tự chấm câu chưa có đáp án. */

/** Chấm lại toàn bộ bài đã nộp theo đáp án hiện có trong sheet CauHoi. */
function regradeAll_() {
  return withLock_(function () {
    var questions = readQuestions_(false);
    var qmap = mapById_(questions);
    var forClass = classQuestionFilter_(questions);
    var sh = getProgSheet_();
    var lr = sh ? sh.getLastRow() : 0;
    if (lr < 2) return { students: 0, cells: 0 };
    var W = progWidth_();
    var data = sh.getRange(2, 1, lr - 1, W).getValues();
    var changed = 0;
    data.forEach(function (vals) {
      var ids = forClass(cellText_(vals[0])).map(function (q) { return q.id; });
      var texts = parseTexts_(vals[COL_TEXTS - 1]);
      var cutSet = {};
      (texts._cut || []).forEach(function (x) { cutSet[x] = 1; });
      ids.forEach(function (id) {
        var idx = qCol_(id) - 1;
        if (cutSet[id]) return;   // chữ đã bị cắt ngắn do quá dài: giữ nguyên kết quả cũ
        var cur = vals[idx];
        if (!cur || cur === ST_REV) return;
        var q = qmap[id];
        var t = str_(texts[id]);
        var nw;
        if (q.answer) nw = statusFor_(q, t);
        else if (cur === ST_SELF_OK || cur === ST_SELF_BAD) nw = cur;
        else nw = t ? ST_PEND : ST_BAD;
        if (nw !== cur) { vals[idx] = nw; changed++; }
      });
      summarizeRow_(vals, ids);
    });
    sh.getRange(2, 1, data.length, W).setValues(data);
    return { students: data.length, cells: changed };
  });
}

// ============================================================================
//  7. DỮ LIỆU CHO BẢNG ĐIỀU KHIỂN
// ============================================================================

function nowIso_() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function dashboard_(req) {
  if (!checkTeacherKey_(req.key)) {
    Utilities.sleep(800);
    return { ok: false, error: 'BAD_KEY' };
  }
  var detail = !!req.detail;           // detail=true: kèm chữ học sinh đã gõ (để xem/xuất từng câu)
  var roster = readRoster_();
  var qs = readQuestions_(false);
  var ids = qs.map(function (q) { return q.id; });
  var sh = getProgSheet_();
  var lr = sh ? sh.getLastRow() : 0;
  var prog = {};
  if (lr >= 2) {
    sh.getRange(2, 1, lr - 1, progWidth_()).getValues().forEach(function (r) {
      prog[fold_(cellText_(r[0])) + '|' + fold_(r[1])] = r;
    });
  }

  var classSet = {}, classes = [];
  var forClass = classQuestionFilter_(qs);
  var visCache = {};
  var visOf = function (cls) {
    var k = classKey_(cls);
    if (!visCache[k]) {
      var m = {}, cnt = 0;
      forClass(cls).forEach(function (q) { m[q.id] = 1; cnt++; });
      visCache[k] = { m: m, n: cnt, asg: ids.map(function (id) { return m[id] ? '1' : '0'; }).join('') };
    }
    return visCache[k];
  };
  var watch = acSummary_();
  var students = roster.map(function (st) {
    if (!classSet[st.cls]) { classSet[st.cls] = 1; classes.push(st.cls); }
    var vis = visOf(st.cls);
    var n = vis.n;
    var r = prog[fold_(st.cls) + '|' + fold_(st.name)];
    var texts = (detail && r) ? parseTexts_(r[COL_TEXTS - 1]) : {};
    var done = 0, good = 0, bad = 0;
    var states = ids.map(function (id) {
      if (!vis.m[id]) return '';
      var v = r ? (r[qCol_(id) - 1] || '') : '';
      if (!v || v === ST_PEND) return '';
      done++;
      if (isGood_(v)) good++; else bad++;
      return v;
    });
    var last = '';
    if (r && r[10] instanceof Date) last = Utilities.formatDate(r[10], TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
    var o = {
      cls: st.cls, name: st.name, loggedIn: !!r,
      done: done, correct: good, wrong: bad, pending: 0,
      score: n ? Math.round(good / n * 100) / 10 : 0,
      attempts: r ? Number(r[9]) || 0 : 0,
      last: last, states: states, n: n, asg: vis.asg
    };
    var w = watch[fold_(st.cls) + '|' + fold_(st.name)];
    var ac = acFor_(st.cls);
    o.acm = ac.mode; o.lim = ac.limit;
    o.lv = w ? w.n : 0; o.lsec = w ? w.secs : 0; o.aus = w ? w.auto : 0;
    if (detail) o.texts = ids.map(function (id) { return str_(texts[id]); });
    return o;
  });
  classes.sort(natCmp_);

  return {
    ok: true, generated: nowIso_(), total: ids.length, classes: classes, students: students,
    questions: qs.map(function (q) {
      return { id: q.id, type: q.type, root: q.root, answer: q.answer, answerFull: q.answer ? answerText_(q) : '', label: qLabel_(q), entry: q.entry, hasKey: !!q.answer, sentence: q.sentence, unit: q.unit, topic: q.topic, scope: q.scope };
    })
  };
}

// ============================================================================
//  7b. HỌC SINH LÀM LẠI TỪ ĐẦU + API CHO TRANG GIÁO VIÊN
// ============================================================================

/** Học sinh bấm "Làm lại từ đầu": xóa kết quả hiện tại của chính em đó (NhatKy vẫn giữ lịch sử). */
function resetMe_(req) {
  var s = authStudent_(req);
  if (!s) return { ok: false, error: 'AUTH' };
  var ids = readQuestionsActive_().filter(function (q) { return inScope_(q.scope, s.cls); }).map(function (q) { return q.id; });
  return withLock_(function () {
    var sh = getProgSheet_();
    var row = findProgressRow_(sh, s.cls, s.name);
    if (!row) return { ok: true };
    var W = progWidth_();
    var vals = sh.getRange(row, 1, 1, W).getValues()[0];
    for (var i = 1; i <= MAX_Q && qCol_(i) <= W; i++) vals[qCol_(i) - 1] = '';
    vals[COL_TEXTS - 1] = '{}';
    summarizeRow_(vals, ids);
    var now = new Date();
    vals[10] = now;
    sh.getRange(row, 1, 1, W).setValues([vals]);
    var log = getSheet_(SH_LOG);
    if (log) log.getRange(log.getLastRow() + 1, 1, 1, 9).setValues([[now, s.cls, s.name, 'Làm lại từ đầu', '', '', '', '', '']]);
    return { ok: true };
  });
}

function teacherOnly_(req, fn) {
  if (!checkTeacherKey_(req.key)) {
    Utilities.sleep(800);
    return { ok: false, error: 'BAD_KEY' };
  }
  return fn();
}

function refreshSummary_() {
  try { lamMoiTongQuan(); } catch (e) { Logger.log('lamMoiTongQuan: ' + e); }
}

// ============================================================================
//  7c. CHỐNG GIAN LẬN: PHÁT HIỆN HỌC SINH RỜI TRANG (v8)
// ============================================================================
//  Cài đặt (Script Properties, khóa AC_CFG):
//    { mode: 'free' | 'exam' | 'all', limit: 4, classes: { <classKey>: { name, mode, limit } } }
//    free = tự do (không theo dõi) | exam = nhắc nhở và ghi nhận khi Làm đề tổng hợp,
//    quá giới hạn thì tự nộp | all = như exam, đồng thời nhắc và ghi nhận cả khi luyện tập.
//  Lớp không có cài đặt riêng dùng cài đặt mặc định. Mặc định ban đầu: free (như bản cũ).
//  Nhật ký ghi ở sheet GiamSat (tự tạo): mỗi lần học sinh quay lại trang là một dòng.

var AC_MODES = ['free', 'exam', 'all'];

function acNormMode_(m) {
  m = String(m == null ? '' : m);
  return AC_MODES.indexOf(m) >= 0 ? m : '';
}

function acNormLimit_(n, dflt) {
  n = parseInt(n, 10);
  if (isNaN(n) || n < 1) return dflt;      // trống / 0 / âm: dùng giá trị mặc định
  return Math.min(20, n);
}

function acRead_() {
  var o = {};
  try { o = JSON.parse(PropertiesService.getScriptProperties().getProperty(AC_PROP) || '{}') || {}; } catch (e) { o = {}; }
  var cfg = { mode: acNormMode_(o.mode) || 'free', limit: acNormLimit_(o.limit, 4), classes: {} };
  var cl = o.classes || {};
  Object.keys(cl).forEach(function (k) {
    var c = cl[k] || {};
    var m = acNormMode_(c.mode), lim = acNormLimit_(c.limit, 0);
    if (m || lim) cfg.classes[k] = { name: String(c.name || k).slice(0, 60), mode: m, limit: lim };
  });
  return cfg;
}

/** Cài đặt hiệu lực cho một lớp: {mode, limit}. */
function acFor_(cls) {
  var cfg = acRead_();
  var ov = cfg.classes[classKey_(cls)];
  return {
    mode: (ov && ov.mode) || cfg.mode,
    limit: (ov && ov.limit) || cfg.limit
  };
}

function acSheet_(create) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SH_WATCH);
  if (!sh && create) {
    sh = ss.insertSheet(SH_WATCH);
    sh.getRange(1, 1, 1, 8).setValues([['Thời gian', 'Lớp', 'Họ và tên', 'Đang ở chế độ', 'Lần rời thứ', 'Thời gian rời (giây)', 'Bài bị tự nộp', 'Ghi chú']]);
    try { sh.setFrozenRows(1); sh.getRange(1, 1, 1, 8).setFontWeight('bold'); } catch (e) { }
  }
  return sh;
}

/** Học sinh báo một lần rời trang (gọi sau khi quay lại). Bỏ qua nếu lớp đang để chế độ tự do. */
function acEvent_(req) {
  var s = authStudent_(req);
  if (!s) return { ok: false, error: 'AUTH' };
  var ac = acFor_(s.cls);
  if (ac.mode === 'free') return { ok: true, ignored: true };
  var evId = String(req.ev || '').slice(0, 60);
  var n = Math.max(0, Math.min(999, parseInt(req.n, 10) || 0));
  var secs = Math.max(0, Math.min(86400, parseInt(req.secs, 10) || 0));
  var where = req.where === 'worksheet' ? 'Làm đề tổng hợp' : 'Luyện tập / sổ tay';
  var auto = req.auto ? 'Có' : '';
  var note = String(req.note || '').replace(/[<>]/g, '').slice(0, 80);
  return withLock_(function () {
    var cache = CacheService.getScriptCache();
    if (evId && cache.get('ac_' + evId)) return { ok: true, dup: true };
    var sh = acSheet_(true);
    sh.getRange(sh.getLastRow() + 1, 1, 1, 8).setValues([[new Date(), s.cls, s.name, where, n, secs, auto, note]]);
    if (evId) cache.put('ac_' + evId, '1', 21600);
    return { ok: true };
  });
}

/** Đọc nhật ký rời trang (tối đa 5000 dòng cuối). */
function acRows_() {
  var sh = acSheet_(false);
  var lr = sh ? sh.getLastRow() : 0;
  if (lr < 2) return [];
  var from = Math.max(2, lr - 5000 + 1);
  return sh.getRange(from, 1, lr - from + 1, 8).getValues().map(function (r) {
    return {
      t: r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, "yyyy-MM-dd'T'HH:mm:ssXXX") : String(r[0] || ''),
      cls: cellText_(r[1]), name: cellText_(r[2]), where: cellText_(r[3]),
      n: Number(r[4]) || 0, secs: Number(r[5]) || 0, auto: String(r[6] || '') === 'Có', note: cellText_(r[7])
    };
  }).filter(function (e) { return e.name; });
}

/** Tổng hợp theo học sinh: { 'lớp|tên': {n, secs, auto, last} } */
function acSummary_() {
  var out = {};
  acRows_().forEach(function (e) {
    var k = fold_(e.cls) + '|' + fold_(e.name);
    var o = out[k] || (out[k] = { n: 0, secs: 0, auto: 0, last: '' });
    o.n++; o.secs += e.secs; if (e.auto) o.auto++;
    if (e.t > o.last) o.last = e.t;
  });
  return out;
}

function acCfgForTeacher_() {
  var cfg = acRead_();
  var classes = listClasses_().classes;
  return {
    mode: cfg.mode, limit: cfg.limit,
    classes: classes.map(function (c) {
      var ov = cfg.classes[classKey_(c)] || {};
      var eff = acFor_(c);
      return { name: c, mode: ov.mode || '', limit: ov.limit || 0, effMode: eff.mode, effLimit: eff.limit };
    })
  };
}

function acList_(req) {
  return teacherOnly_(req, function () {
    var rows = acRows_();
    var recent = rows.slice(-1000).reverse();
    return { ok: true, cfg: acCfgForTeacher_(), events: recent };
  });
}

function acSave_(req) {
  return teacherOnly_(req, function () {
    var c = req.cfg || {};
    var mode = acNormMode_(c.mode) || 'free';
    var limit = acNormLimit_(c.limit, 4);
    var classes = {};
    (Array.isArray(c.classes) ? c.classes : []).forEach(function (x) {
      if (!x || !x.name) return;
      var m = acNormMode_(x.mode), lim = acNormLimit_(x.limit, 0);
      if (m || lim) classes[classKey_(x.name)] = { name: String(x.name).slice(0, 60), mode: m, limit: lim };
    });
    PropertiesService.getScriptProperties().setProperty(AC_PROP, JSON.stringify({ mode: mode, limit: limit, classes: classes }));
    return { ok: true, cfg: acCfgForTeacher_() };
  });
}


// ---------- Câu hỏi ----------

function tqListData_() {
  var fams = readFamilies_();
  var names = {};
  fams.forEach(function (f) { names[fold_(f.name)] = 1; });
  var entries = readEntries_();
  var entNames = {};
  entries.forEach(function (e) { entNames[fold_(e.name)] = 1; });
  var qs = readQuestions_(true, true).map(function (q) {
    return {
      id: q.id, type: q.type, sentence: q.sentence, root: q.root, answer: q.answer, alts: q.alts.join(' | '),
      explain: q.explain, translation: q.translation, family: q.family, hidden: q.hidden,
      opts: q.opts, entry: q.entry, unit: q.unit, topic: q.topic, scope: q.scope,
      problem: q.typeBad ? 'Dạng bài không hợp lệ' : questionProblem_(q, names, entNames)
    };
  });
  return { questions: qs, families: fams, entries: entries, maxQ: MAX_Q, unitSuggest: UNIT_SUGGEST, classes: listClasses_().classes };
}

function tqList_(req) {
  return teacherOnly_(req, function () {
    var d = tqListData_();
    d.ok = true;
    return d;
  });
}

function cleanText_(v, max) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
}

/** Kiểm tra và chuẩn hóa một câu hỏi do giáo viên nhập. Trả {v} hoặc {error, message}. */
function validateQuestion_(q) {
  var tp = typeOf_(q.type);
  if (!tp) return { error: 'BAD_TYPE', message: 'Dạng bài không hợp lệ (chọn: Điền từ, Trắc nghiệm, Viết lại câu, Word entry).' };
  var sentence = cleanText_(q.sentence, 500);
  var root = cleanText_(q.root, 120);
  var answer = cleanText_(q.answer, 250);
  var explain = cleanText_(q.explain, 1500);
  var translation = cleanText_(q.translation, 800);
  var family = cleanText_(q.family, 60);
  var entry = cleanText_(q.entry, 60);
  var unit = cleanText_(q.unit, 60);
  var topic = cleanText_(q.topic, 60);
  var scope = cleanScope_(q.scope);
  var opts = ['', '', '', ''];
  var rawOpts = Array.isArray(q.opts) ? q.opts : [q.a, q.b, q.c, q.d];
  for (var i = 0; i < 4; i++) opts[i] = cleanText_(rawOpts[i], 200);
  var all = [sentence, root, answer, explain, translation, family, entry, unit, topic, opts.join(''), String(q.alts || ''), String(q.scope == null ? '' : q.scope)].join('');
  if (/[<>]/.test(all)) return { error: 'BAD_CHAR', message: 'Không dùng ký tự < hoặc > trong câu hỏi.' };
  var blanks = (sentence.match(/_{3,}/g) || []).length;
  if (!sentence) return { error: 'NO_SENTENCE', message: 'Chưa nhập câu hỏi.' };

  if (tp === 'trac_nghiem') {
    if (!answer) return { error: 'NO_ANSWER', message: 'Chưa nhập đáp án.' };
    var filled = opts.filter(function (o) { return o; }).length;
    if (filled < 2) return { error: 'FEW_OPTIONS', message: 'Trắc nghiệm cần ít nhất 2 lựa chọn (A, B, C, D).' };
    var L = letterOf_(answer);
    if (!L) return { error: 'BAD_ANSWER', message: 'Đáp án trắc nghiệm phải là một chữ cái: A, B, C hoặc D.' };
    if (!opts[OPT_LETTERS.indexOf(L)]) return { error: 'BAD_ANSWER', message: 'Đáp án ' + L + ' đang trỏ tới lựa chọn để trống.' };
    answer = L;
  } else {
    opts = ['', '', '', ''];
    if (!answer && (tp === 'viet_lai')) return { error: 'NO_ANSWER', message: 'Chưa nhập đáp án.' };
    if (/[|;]/.test(answer)) return { error: 'BAD_ANSWER', message: 'Đáp án chính không chứa dấu | hoặc ;. Đáp án khác nhập ở ô "Đáp án khác".' };
    if (tp === 'dien_tu' || tp === 'word_entry') {
      if (blanks === 0) return { error: 'NO_BLANK', message: 'Câu hỏi cần có 1 chỗ trống: gõ ____ vào chỗ học sinh phải điền.' };
      if (blanks > 1) return { error: 'MANY_BLANKS', message: 'Câu hỏi chỉ được có 1 chỗ trống.' };
    }
    if (tp === 'dien_tu' && !root) return { error: 'NO_ROOT', message: 'Chưa nhập từ gốc (từ cho trong ngoặc).' };
    if (!answer) return { error: 'NO_ANSWER', message: 'Chưa nhập đáp án.' };
    if (tp === 'word_entry' && !entry) return { error: 'NO_ENTRY', message: 'Chưa chọn mục từ điển cho Word entry.' };
  }
  if (tp !== 'word_entry') entry = '';
  var alts = [], seenAlt = {};
  seenAlt[answer.toLowerCase()] = 1;
  if (tp !== 'trac_nghiem') {
    splitAlts_(q.alts).forEach(function (a) {
      a = cleanText_(a, 250);
      if (a && !seenAlt[a.toLowerCase()]) { seenAlt[a.toLowerCase()] = 1; alts.push(a); }
    });
  }
  return {
    v: {
      type: tp,
      sentence: (tp === 'viet_lai') ? sentence : sentence.replace(/_{3,}/g, '__________'),
      root: root, answer: answer, alts: alts,
      explain: explain, translation: translation, family: tp === 'dien_tu' ? family : '', hidden: !!q.hidden,
      opts: opts, entry: entry, unit: unit, topic: topic, scope: scope
    }
  };
}

/** Số STT còn trống nhỏ nhất (1..MAX_Q) trong tập 'taken'; 0 nếu hết chỗ. */
function lowestFree_(taken) {
  for (var k = 1; k <= MAX_Q; k++) { if (!taken[k]) return k; }
  return 0;
}

function questionIndex_() {
  var sh = getQSheet_();
  var lr = sh ? sh.getLastRow() : 0;
  var rowOf = {}, taken = {};
  if (lr >= 2) {
    sh.getRange(2, 1, lr - 1, 1).getValues().forEach(function (r, i) {
      var id = parseInt(r[0], 10);
      if (id >= 1 && id <= MAX_Q) { taken[id] = 1; if (!rowOf[id]) rowOf[id] = i + 2; }
    });
  }
  return { sh: sh, lastRow: lr, rowOf: rowOf, taken: taken };
}

/** Một dòng 18 ô của sheet CauHoi từ câu hỏi đã chuẩn hóa. */
function questionRow_(id, v) {
  return [id, v.sentence, v.root, v.answer, v.alts.join('|'), v.explain, v.translation, v.family,
    v.hidden ? 'x' : '', TYPE_LABEL[v.type], v.opts[0], v.opts[1], v.opts[2], v.opts[3], v.entry, v.unit || '', v.topic || '', v.scope || ''];
}

function writeQuestionRow_(sh, row, id, v) {
  sh.getRange(row, 2, 1, Q_COLS - 1).setNumberFormat('@');
  sh.getRange(row, 1, 1, Q_COLS).setValues([questionRow_(id, v)]);
}

function entryExists_(name) {
  var f = fold_(name);
  return readEntries_().some(function (x) { return fold_(x.name) === f; });
}

function familyExists_(name) {
  var f = fold_(name);
  return readFamilies_().some(function (x) { return fold_(x.name) === f; });
}

function tqSave_(req) {
  return teacherOnly_(req, function () {
    var q = req.q || {};
    var chk = validateQuestion_(q);
    if (chk.error) return { ok: false, error: chk.error, message: chk.message };
    var v = chk.v;
    if (v.family && !familyExists_(v.family)) {
      return { ok: false, error: 'BAD_FAMILY', message: 'Gia đình từ "' + v.family + '" chưa có. Hãy thêm ở thẻ Gia đình từ trước, hoặc để trống.' };
    }
    if (v.entry && !entryExists_(v.entry)) {
      return { ok: false, error: 'BAD_ENTRY', message: 'Mục từ điển "' + v.entry + '" chưa có. Hãy thêm mục từ điển trước.' };
    }
    var res = withLock_(function () {
      var ix = questionIndex_();
      if (!ix.sh) return { ok: false, error: 'NO_SHEET', message: 'Chưa thiết lập: chạy menu Word Form 9 > 1. Thiết lập ban đầu.' };
      var id = parseInt(q.id, 10), row;
      if (id >= 1 && id <= MAX_Q && ix.rowOf[id]) {
        row = ix.rowOf[id];
      } else {
        if (!(id >= 1 && id <= MAX_Q) || ix.taken[id]) id = lowestFree_(ix.taken);
        if (!id) return { ok: false, error: 'FULL', message: 'Đã đủ ' + MAX_Q + ' câu.' };
        row = Math.max(ix.lastRow, 1) + 1;
      }
      writeQuestionRow_(ix.sh, row, id, v);
      return { ok: true, id: id };
    });
    if (!res.ok) return res;
    refreshSummary_();
    var d = tqListData_(); d.ok = true; d.id = res.id;
    return d;
  });
}

function tqDelete_(req) {
  return teacherOnly_(req, function () {
    var id = parseInt(req.id, 10);
    var res = withLock_(function () {
      var ix = questionIndex_();
      if (!ix.rowOf[id]) return { ok: false, error: 'NOT_FOUND', message: 'Không tìm thấy câu này.' };
      ix.sh.deleteRows(ix.rowOf[id], 1);
      return { ok: true };
    });
    if (!res.ok) return res;
    refreshSummary_();
    var d = tqListData_(); d.ok = true;
    return d;
  });
}

/**
 * Gán / xóa Unit, Chủ đề và "Áp dụng cho" (khối/lớp) cho nhiều câu cùng lúc (ô để trống thì giữ nguyên, trừ khi chọn xóa).
 * scopeMode: set (thay bằng) | add (thêm vào) | remove (bớt đi) | clear (về "mọi lớp"). Bỏ trống = không đổi.
 */
function tqTag_(req) {
  return teacherOnly_(req, function () {
    var want = {}, n = 0;
    (Array.isArray(req.ids) ? req.ids : []).forEach(function (x) {
      var id = parseInt(x, 10);
      if (id >= 1 && id <= MAX_Q && !want[id]) { want[id] = 1; n++; }
    });
    if (!n) return { ok: false, error: 'NO_IDS', message: 'Chưa chọn câu nào.' };
    var rawU = String(req.unit == null ? '' : req.unit), rawT = String(req.topic == null ? '' : req.topic);
    var rawS = String(req.scope == null ? '' : req.scope);
    if (/[<>]/.test(rawU + rawT + rawS)) return { ok: false, error: 'BAD_CHAR', message: 'Không dùng ký tự < hoặc > trong Unit / Chủ đề / Áp dụng cho.' };
    var u = cleanLabel_(rawU), t = cleanLabel_(rawT), sc = cleanScope_(rawS);
    var clrU = !!req.clearUnit, clrT = !!req.clearTopic;
    var mode = String(req.scopeMode || '');
    if (mode && mode !== 'set' && mode !== 'add' && mode !== 'remove' && mode !== 'clear') mode = '';
    if ((mode === 'set' || mode === 'add' || mode === 'remove') && !scopeTokens_(sc).length) {
      return { ok: false, error: 'NO_SCOPE', message: 'Hãy nhập khối hoặc lớp (ví dụ: 7 hoặc 6|7 hoặc 9/1).' };
    }
    if (!u && !t && !clrU && !clrT && !mode) return { ok: false, error: 'NOTHING', message: 'Hãy nhập Unit, Chủ đề hoặc khối/lớp cần gán (hoặc chọn xóa).' };
    var res = withLock_(function () {
      var sh = getQSheet_();
      if (!sh) return { ok: false, error: 'NO_SHEET', message: 'Chưa thiết lập.' };
      var lr = sh.getLastRow();
      if (lr < 2) return { ok: false, error: 'NOT_FOUND', message: 'Chưa có câu hỏi.' };
      var ids = sh.getRange(2, 1, lr - 1, 1).getValues();
      var cur = sh.getRange(2, 16, lr - 1, 3).getValues();   // P Unit, Q Chủ đề, R Áp dụng cho
      var count = 0, nowAll = 0;
      for (var i = 0; i < ids.length; i++) {
        var id = parseInt(ids[i][0], 10);
        if (!want[id]) continue;
        if (clrU) cur[i][0] = ''; else if (u) cur[i][0] = u;
        if (clrT) cur[i][1] = ''; else if (t) cur[i][1] = t;
        if (mode) {
          cur[i][2] = mergeScope_(cleanScope_(cur[i][2]), sc, mode);
          if (!cur[i][2]) nowAll++;
        }
        count++;
      }
      if (!count) return { ok: false, error: 'NOT_FOUND', message: 'Không tìm thấy các câu đã chọn.' };
      sh.getRange(2, 16, cur.length, 3).setNumberFormat('@');
      sh.getRange(2, 16, cur.length, 3).setValues(cur);
      return { ok: true, count: count, nowAll: mode ? nowAll : 0 };
    });
    if (!res.ok) return res;
    var d = tqListData_(); d.ok = true; d.count = res.count; d.nowAll = res.nowAll;
    return d;
  });
}

/**
 * Dán nhiều câu từ Excel/Word. Mỗi dòng, các cột cách nhau bằng Tab, đúng thứ tự cột của sheet CauHoi:
 *   Câu hỏi | Từ gốc / Gợi ý | Đáp án | Đáp án khác | Giải thích | Dịch nghĩa | Gia đình từ | Ẩn | Dạng bài | A | B | C | D | Mục từ điển | Unit | Chủ đề | Áp dụng cho
 * (có thể có thêm cột STT ở đầu, có thể bỏ các cột cuối; dòng tiêu đề tự bỏ qua).
 */
function parseQuestionLines_(text) {
  var out = [];
  String(text == null ? '' : text).split(/\r?\n/).forEach(function (line, i) {
    if (!line.replace(/[\s\t]/g, '')) return;
    var c = line.split('\t').map(function (x) { return x.trim(); });
    if (/^\d+$/.test(c[0]) && c.length > 2) { c.shift(); }
    if (i === 0 && !/_{3,}/.test(c[0] || '') && /câu|sentence|question|stt/i.test(c[0] || '')) return; // dòng tiêu đề
    out.push({ line: i + 1, q: {
      sentence: c[0], root: c[1], answer: c[2], alts: c[3], explain: c[4], translation: c[5], family: c[6],
      hidden: !!str_(c[7]), type: c[8], opts: [c[9], c[10], c[11], c[12]], entry: c[13], unit: c[14], topic: c[15], scope: c[16]
    } });
  });
  return out;
}

function tqImport_(req) {
  return teacherOnly_(req, function () {
    var parsed = parseQuestionLines_(req.text);
    if (!parsed.length) return { ok: false, error: 'EMPTY', message: 'Chưa có dòng nào để thêm.' };
    var defScope = cleanScope_(req.scope);   // khối/lớp áp cho các dòng không có cột "Áp dụng cho"
    var res = withLock_(function () {
      var ix = questionIndex_();
      if (!ix.sh) return { ok: false, error: 'NO_SHEET', message: 'Chưa thiết lập: chạy menu Word Form 9 > 1. Thiết lập ban đầu.' };
      var have = {};
      readQuestions_(true, true).forEach(function (x) { have[fold_(x.sentence)] = 1; });
      var rows = [], skipped = [];
      parsed.forEach(function (p) {
        var chk = validateQuestion_(p.q);
        if (chk.error) { skipped.push({ line: p.line, reason: chk.message }); return; }
        var v = chk.v;
        if (!v.scope && defScope) v.scope = defScope;
        if (v.family && !familyExists_(v.family)) {
          skipped.push({ line: p.line, reason: 'Gia đình từ "' + v.family + '" chưa có (thêm ở thẻ Gia đình từ trước).' });
          return;
        }
        if (v.entry && !entryExists_(v.entry)) {
          skipped.push({ line: p.line, reason: 'Mục từ điển "' + v.entry + '" chưa có (thêm ở thẻ Từ điển trước).' });
          return;
        }
        if (have[fold_(v.sentence)]) { skipped.push({ line: p.line, reason: 'Trùng với câu đã có.' }); return; }
        var id = lowestFree_(ix.taken);
        if (!id) { skipped.push({ line: p.line, reason: 'Đã đủ ' + MAX_Q + ' câu.' }); return; }
        ix.taken[id] = 1; have[fold_(v.sentence)] = 1;
        rows.push(questionRow_(id, v));
      });
      if (rows.length) {
        var start = Math.max(ix.lastRow, 1) + 1;
        ix.sh.getRange(start, 2, rows.length, Q_COLS - 1).setNumberFormat('@');
        ix.sh.getRange(start, 1, rows.length, Q_COLS).setValues(rows);
      }
      return { ok: true, added: rows.length, skipped: skipped };
    });
    if (!res.ok) return res;
    if (res.added) refreshSummary_();
    var d = tqListData_(); d.ok = true; d.added = res.added; d.skipped = res.skipped;
    return d;
  });
}

/** Các dòng câu hỏi mẫu (17 ô) theo thứ tự cột của sheet CauHoi: 40 câu điền từ của file gốc + các dạng bài mới. */
function sampleRows_() {
  var rows = SAMPLE_QUESTIONS.map(function (q, i) {
    return q.concat(['', TYPE_LABEL.dien_tu, '', '', '', '', '']).concat(SAMPLE_META_OLD[i] || ['', '']);
  });
  return rows.concat(SAMPLE_QUESTIONS_NEW).map(padRow_);
}

/** Bổ sung ô trống cho đủ Q_COLS cột (câu mẫu cũ chỉ có 17 ô). */
function padRow_(r) {
  var o = r.slice();
  while (o.length < Q_COLS) o.push('');
  return o;
}

/** Nạp lại bộ câu hỏi mẫu (40 câu điền từ của file gốc + các dạng mới), chỉ khi chưa có câu hỏi nào. */
function tqSeed_(req) {
  return teacherOnly_(req, function () {
    var res = withLock_(function () {
      var ix = questionIndex_();
      if (!ix.sh) return { ok: false, error: 'NO_SHEET', message: 'Chưa thiết lập: chạy menu Word Form 9 > 1. Thiết lập ban đầu.' };
      if (ix.lastRow >= 2) return { ok: false, error: 'NOT_EMPTY', message: 'Đã có câu hỏi, không nạp lại để khỏi ghi đè.' };
      var fa = getSheet_(SH_FAMILY);
      if (fa && fa.getLastRow() < 2) fa.getRange(2, 1, SAMPLE_FAMILIES.length, 2).setValues(SAMPLE_FAMILIES);
      var en = getSheet_(SH_ENTRY);
      if (en && en.getLastRow() < 2) en.getRange(2, 1, SAMPLE_ENTRIES.length, 5).setValues(SAMPLE_ENTRIES);
      var rows = sampleRows_();
      ix.sh.getRange(2, 2, rows.length, Q_COLS - 1).setNumberFormat('@');
      ix.sh.getRange(2, 1, rows.length, Q_COLS).setValues(rows);
      return { ok: true };
    });
    if (!res.ok) return res;
    refreshSummary_();
    var d = tqListData_(); d.ok = true;
    return d;
  });
}

// ---------- Gia đình từ ----------

function tfSave_(req) {
  return teacherOnly_(req, function () {
    var name = cleanText_(req.name, 60), oldName = cleanText_(req.oldName, 60);
    var content = cleanText_(req.content, 1500);
    if (/[<>]/.test(name + content)) return { ok: false, error: 'BAD_CHAR', message: 'Không dùng ký tự < hoặc >.' };
    if (!name) return { ok: false, error: 'NO_NAME', message: 'Chưa nhập từ gốc (tên nhóm).' };
    if (!content) return { ok: false, error: 'NO_CONTENT', message: 'Chưa nhập nội dung gia đình từ.' };
    var res = withLock_(function () {
      var sh = getSheet_(SH_FAMILY);
      if (!sh) return { ok: false, error: 'NO_SHEET', message: 'Chưa thiết lập.' };
      var lr = sh.getLastRow();
      var rows = lr >= 2 ? sh.getRange(2, 1, lr - 1, 1).getValues() : [];
      var find = function (n) {
        var f = fold_(n);
        for (var i = 0; i < rows.length; i++) { if (fold_(rows[i][0]) === f) return i + 2; }
        return 0;
      };
      var target = oldName ? find(oldName) : find(name);
      var clash = find(name);
      if (clash && (clash !== target || !oldName)) return { ok: false, error: 'DUP_NAME', message: 'Đã có nhóm "' + name + '".' };
      var renamed = false;
      if (target) {
        sh.getRange(target, 1, 1, 2).setNumberFormat('@');
        sh.getRange(target, 1, 1, 2).setValues([[name, content]]);
        renamed = oldName && fold_(oldName) !== fold_(name);
      } else {
        var row = Math.max(lr, 1) + 1;
        sh.getRange(row, 1, 1, 2).setNumberFormat('@');
        sh.getRange(row, 1, 1, 2).setValues([[name, content]]);
      }
      if (renamed) {
        // đổi tên nhóm -> cập nhật các câu đang dùng nhóm đó
        var qs = getSheet_(SH_QUESTIONS), ql = qs ? qs.getLastRow() : 0;
        if (ql >= 2) {
          var col = qs.getRange(2, 8, ql - 1, 1).getValues();
          var changed = false;
          col.forEach(function (r) { if (fold_(r[0]) === fold_(oldName)) { r[0] = name; changed = true; } });
          if (changed) qs.getRange(2, 8, ql - 1, 1).setValues(col);
        }
      }
      return { ok: true };
    });
    if (!res.ok) return res;
    var d = tqListData_(); d.ok = true;
    return d;
  });
}

function tfDelete_(req) {
  return teacherOnly_(req, function () {
    var name = cleanText_(req.name, 60);
    var res = withLock_(function () {
      var sh = getSheet_(SH_FAMILY);
      var lr = sh ? sh.getLastRow() : 0;
      var rows = lr >= 2 ? sh.getRange(2, 1, lr - 1, 1).getValues() : [];
      var target = 0;
      for (var i = 0; i < rows.length; i++) { if (fold_(rows[i][0]) === fold_(name)) { target = i + 2; break; } }
      if (!target) return { ok: false, error: 'NOT_FOUND', message: 'Không tìm thấy nhóm này.' };
      sh.deleteRows(target, 1);
      var qs = getSheet_(SH_QUESTIONS), ql = qs ? qs.getLastRow() : 0;
      if (ql >= 2) {
        var col = qs.getRange(2, 8, ql - 1, 1).getValues();
        var changed = false;
        col.forEach(function (r) { if (fold_(r[0]) === fold_(name)) { r[0] = ''; changed = true; } });
        if (changed) qs.getRange(2, 8, ql - 1, 1).setValues(col);
      }
      return { ok: true };
    });
    if (!res.ok) return res;
    var d = tqListData_(); d.ok = true;
    return d;
  });
}

// ---------- Mục từ điển (Word entry) ----------

function teSave_(req) {
  return teacherOnly_(req, function () {
    var name = cleanText_(req.name, 60), oldName = cleanText_(req.oldName, 60);
    var ipa = cleanText_(req.ipa, 80), pos = cleanText_(req.pos, 40);
    var defs = splitLines_(req.defs).map(function (x) { return cleanText_(x, 300); }).filter(Boolean);
    var examples = splitLines_(req.examples).map(function (x) { return cleanText_(x, 300); }).filter(Boolean);
    if (/[<>]/.test(name + ipa + pos + defs.join('') + examples.join(''))) return { ok: false, error: 'BAD_CHAR', message: 'Không dùng ký tự < hoặc >.' };
    if (!name) return { ok: false, error: 'NO_NAME', message: 'Chưa nhập từ cần tra (headword).' };
    if (!defs.length) return { ok: false, error: 'NO_DEFS', message: 'Chưa nhập nghĩa của từ.' };
    if (examples.length < 2) return { ok: false, error: 'FEW_EXAMPLES', message: 'Cần ít nhất 2 câu ví dụ (cụm cần học viết giữa hai dấu **, ví dụ: **make a decision**).' };
    var res = withLock_(function () {
      var sh = getSheet_(SH_ENTRY);
      if (!sh) return { ok: false, error: 'NO_SHEET', message: 'Chưa có sheet TuDien: chạy menu Word Form 9 > 8. Nâng cấp lên 4 dạng bài.' };
      var lr = sh.getLastRow();
      var rows = lr >= 2 ? sh.getRange(2, 1, lr - 1, 1).getValues() : [];
      var find = function (n) {
        var f = fold_(n);
        for (var i = 0; i < rows.length; i++) { if (fold_(rows[i][0]) === f) return i + 2; }
        return 0;
      };
      var target = oldName ? find(oldName) : find(name);
      var clash = find(name);
      if (clash && (clash !== target || !oldName)) return { ok: false, error: 'DUP_NAME', message: 'Đã có mục từ "' + name + '".' };
      var vals = [[name, ipa, pos, defs.join('\n'), examples.join('\n')]];
      var row = target || (Math.max(lr, 1) + 1);
      sh.getRange(row, 1, 1, 5).setNumberFormat('@');
      sh.getRange(row, 1, 1, 5).setValues(vals);
      if (target && oldName && fold_(oldName) !== fold_(name)) {
        var qs = getSheet_(SH_QUESTIONS), ql = qs ? qs.getLastRow() : 0;
        if (ql >= 2) {
          var col = qs.getRange(2, 15, ql - 1, 1).getValues(), changed = false;
          col.forEach(function (r) { if (fold_(r[0]) === fold_(oldName)) { r[0] = name; changed = true; } });
          if (changed) qs.getRange(2, 15, ql - 1, 1).setValues(col);
        }
      }
      return { ok: true };
    });
    if (!res.ok) return res;
    var d = tqListData_(); d.ok = true;
    return d;
  });
}

function teDelete_(req) {
  return teacherOnly_(req, function () {
    var name = cleanText_(req.name, 60);
    var used = readQuestions_(true, true).filter(function (q) { return q.type === 'word_entry' && fold_(q.entry) === fold_(name); });
    if (used.length) {
      return { ok: false, error: 'IN_USE', message: 'Mục từ này đang được dùng ở câu ' + used.map(function (q) { return q.id; }).join(', ') + '. Hãy sửa hoặc xóa các câu đó trước.' };
    }
    var res = withLock_(function () {
      var sh = getSheet_(SH_ENTRY);
      var lr = sh ? sh.getLastRow() : 0;
      var rows = lr >= 2 ? sh.getRange(2, 1, lr - 1, 1).getValues() : [];
      for (var i = 0; i < rows.length; i++) {
        if (fold_(rows[i][0]) === fold_(name)) { sh.deleteRows(i + 2, 1); return { ok: true }; }
      }
      return { ok: false, error: 'NOT_FOUND', message: 'Không tìm thấy mục từ này.' };
    });
    if (!res.ok) return res;
    var d = tqListData_(); d.ok = true;
    return d;
  });
}

// ---------- Danh sách lớp ----------

function rosterData_() {
  var students = readRoster_();
  var classes = [], seen = {};
  students.forEach(function (s) { if (!seen[s.cls]) { seen[s.cls] = 1; classes.push(s.cls); } });
  classes.sort(natCmp_);
  return { students: students, classes: classes };
}

function trList_(req) {
  return teacherOnly_(req, function () {
    var d = rosterData_(); d.ok = true;
    return d;
  });
}

function trImport_(req) {
  return teacherOnly_(req, function () {
    var r = withLock_(function () { return importRoster(req.cls, req.text); });
    var d = rosterData_(); d.ok = true;
    d.added = r.added; d.skipped = r.skipped; d.bad = r.bad;
    return d;
  });
}

function rosterRowOf_(sh, cls, name) {
  var lr = sh.getLastRow();
  if (lr < 2) return 0;
  var rows = sh.getRange(2, 1, lr - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (fold_(cellText_(rows[i][0])) === fold_(cls) && fold_(cellText_(rows[i][1])) === fold_(name)) return i + 2;
  }
  return 0;
}

function trDelete_(req) {
  return teacherOnly_(req, function () {
    var res = withLock_(function () {
      var sh = getSheet_(SH_ROSTER);
      var row = sh ? rosterRowOf_(sh, req.cls, req.name) : 0;
      if (!row) return { ok: false, error: 'NOT_FOUND', message: 'Không tìm thấy học sinh này.' };
      sh.deleteRows(row, 1);
      return { ok: true };
    });
    if (!res.ok) return res;
    var d = rosterData_(); d.ok = true;
    return d;
  });
}

function trCode_(req) {
  return teacherOnly_(req, function () {
    var code = cleanText_(req.code, 30);
    var res = withLock_(function () {
      var sh = getSheet_(SH_ROSTER);
      var row = sh ? rosterRowOf_(sh, req.cls, req.name) : 0;
      if (!row) return { ok: false, error: 'NOT_FOUND', message: 'Không tìm thấy học sinh này.' };
      sh.getRange(row, 3).setNumberFormat('@');
      sh.getRange(row, 3).setValue(code);
      return { ok: true };
    });
    if (!res.ok) return res;
    var d = rosterData_(); d.ok = true;
    return d;
  });
}

/** Chấm lại toàn bộ bài đã nộp theo đáp án hiện tại (sau khi giáo viên sửa đáp án). */
function regradeApi_(req) {
  return teacherOnly_(req, function () {
    var r = regradeAll_();
    refreshSummary_();
    return { ok: true, students: r.students, cells: r.cells };
  });
}

// ============================================================================
//  8. NHẬP NHANH DANH SÁCH LỚP
// ============================================================================

function titleCaseIfShout_(name) {
  // Chữ IN HOA toàn bộ hoặc chữ thường toàn bộ -> Viết Hoa Chữ Cái Đầu Mỗi Từ
  var hasLetter = name !== name.toUpperCase() || name !== name.toLowerCase();
  if (hasLetter && (name === name.toUpperCase() || name === name.toLowerCase())) {
    return name.toLowerCase().split(' ').map(function (w) {
      return w ? w.charAt(0).toUpperCase() + w.slice(1) : w;
    }).join(' ');
  }
  return name;
}

/** Phân tích văn bản dán vào: trả {rows:[{cls,name}], bad:[dòng lỗi]} (hàm thuần). */
function parseRosterText_(cls, text) {
  var rows = [], bad = [];
  cls = cellText_(cls);
  String(text == null ? '' : text).split(/\r?\n/).forEach(function (line) {
    line = line.replace(/ /g, ' ').trim();
    if (!line) return;
    var c = cls, name = '';
    if (line.indexOf('\t') >= 0) {
      var cols = line.split('\t').map(function (x) { return x.trim(); }).filter(function (x) { return x; });
      var isNameCol = function (x) { return /[A-Za-zÀ-ỹ]/.test(x) && !/^\d+$/.test(x); };
      if (!c) {
        // định dạng: Lớp <tab> Họ tên
        if (cols.length >= 2 && cols[0].length <= 10) {
          c = cols[0];
          var rest = cols.slice(1).filter(isNameCol);
          name = rest[0] || '';
        }
      } else {
        var cand = cols.filter(isNameCol);
        name = cand[0] || '';
      }
    } else {
      name = line.replace(/^\d+\s*[.)\-:]\s*/, '').replace(/^\d+\s+(?=[A-Za-zÀ-ỹ])/, '');
      if (!c && /[,;]/.test(name)) {
        var p = name.split(/[,;]/).map(function (x) { return x.trim(); });
        if (p.length >= 2 && p[0].length <= 10) { c = p[0]; name = p[1]; }
      }
    }
    name = String(name).replace(/\s+/g, ' ').trim();
    if (!c || name.length < 2 || !/[A-Za-zÀ-ỹ]/.test(name)) { bad.push(line); return; }
    rows.push({ cls: c, name: titleCaseIfShout_(name) });
  });
  return { rows: rows, bad: bad };
}

/** Gọi từ hộp thoại nhập nhanh. Thêm vào cuối DanhSach, bỏ qua học sinh đã có. */
function importRoster(cls, text) {
  var parsed = parseRosterText_(cls, text);
  var sh = getSheet_(SH_ROSTER);
  if (!sh) throw new Error('Chưa có sheet DanhSach: hãy chạy 1. Thiết lập ban đầu');
  var have = {};
  readRoster_().forEach(function (r) { have[fold_(r.cls) + '|' + fold_(r.name)] = 1; });
  var add = [], skipped = [];
  parsed.rows.forEach(function (r) {
    var k = fold_(r.cls) + '|' + fold_(r.name);
    if (have[k]) { skipped.push(r.name); return; }
    have[k] = 1;
    add.push([r.cls, r.name, '', '']);
  });
  if (add.length) {
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, add.length, 3).setNumberFormat('@');
    sh.getRange(start, 1, add.length, 4).setValues(add);
  }
  return { added: add.length, skipped: skipped, bad: parsed.bad };
}

var ROSTER_DIALOG_HTML =
  '<div style="font-family:Arial,sans-serif;font-size:14px;padding:6px 10px;">' +
  '<p><b>Cách 1:</b> gõ tên lớp (ví dụ 9A1), rồi dán danh sách họ tên, mỗi dòng một em. Số thứ tự đầu dòng tự bỏ; chữ IN HOA hoặc chữ thường toàn bộ tự đổi thành Viết Hoa Đầu Từ.</p>' +
  '<p><b>Cách 2:</b> để trống ô Lớp, dán 2 cột <i>Lớp ⇥ Họ tên</i> (copy từ Excel).</p>' +
  '<input id="cls" placeholder="Lớp, ví dụ 9A1" style="width:160px;padding:6px;margin-bottom:8px;"><br>' +
  '<textarea id="txt" rows="13" style="width:100%;box-sizing:border-box;padding:6px;" placeholder="Nguyễn Văn An&#10;Lê Thị Hoa&#10;..."></textarea>' +
  '<p><button id="go" onclick="go()" style="padding:8px 16px;font-weight:bold;">Thêm vào danh sách</button> ' +
  '<button onclick="google.script.host.close()" style="padding:8px 16px;">Đóng</button></p>' +
  '<div id="msg" style="white-space:pre-wrap;"></div></div>' +
  '<script>function go(){var b=document.getElementById("go");b.disabled=true;' +
  'google.script.run.withSuccessHandler(function(r){b.disabled=false;var m="Đã thêm: "+r.added+" học sinh.";' +
  'if(r.skipped.length)m+="\\nĐã có sẵn, bỏ qua: "+r.skipped.length;' +
  'if(r.bad.length)m+="\\nKhông đọc được "+r.bad.length+" dòng:\\n"+r.bad.slice(0,8).join("\\n");' +
  'document.getElementById("msg").innerText=m;if(r.added)document.getElementById("txt").value="";})' +
  '.withFailureHandler(function(e){b.disabled=false;document.getElementById("msg").innerText="Lỗi: "+e.message;})' +
  '.importRoster(document.getElementById("cls").value,document.getElementById("txt").value);}</script>';

function nhapDanhSachNhanh() {
  var html = HtmlService.createHtmlOutput(ROSTER_DIALOG_HTML).setWidth(560).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Nhập nhanh danh sách lớp');
}

// ============================================================================
//  9. KIỂM TRA DỮ LIỆU (DANH SÁCH + CÂU HỎI)
// ============================================================================

function listShort_(arr, max) {
  max = max || 10;
  return arr.slice(0, max).join(', ') + (arr.length > max ? ', ... (+' + (arr.length - max) + ')' : '');
}

/** Kiểm tra và tự đánh số STT. Trả về {text, problems, filled}. */
function checkData_() {
  var lines = [], problems = 0, filled = 0;

  // --- Danh sách lớp ---
  var ro = getSheet_(SH_ROSTER);
  var rLast = ro ? ro.getLastRow() : 0;
  var rawRoster = rLast >= 2 ? ro.getRange(2, 1, rLast - 1, 3).getValues() : [];
  var count = {}, variants = {}, dupes = [], seen = {}, partial = [], digits = [];
  var total = 0;
  rawRoster.forEach(function (r, i) {
    var c = cellText_(r[0]), n = cellText_(r[1]);
    if (!c && !n) return;
    if (!c || !n) { partial.push('dòng ' + (i + 2)); return; }
    total++;
    count[c] = (count[c] || 0) + 1;
    var f = fold_(c);
    variants[f] = variants[f] || {};
    variants[f][c] = 1;
    var k = f + '|' + fold_(n);
    if (seen[k]) dupes.push(c + ' - ' + n); else seen[k] = 1;
    if (/\d/.test(n)) digits.push(n);
  });
  lines.push('DANH SÁCH LỚP: ' + total + ' học sinh, ' + Object.keys(count).length + ' lớp');
  if (!total) { lines.push('  ✗ Chưa có học sinh nào (dùng menu 3. Nhập nhanh danh sách).'); problems++; }
  if (partial.length) { lines.push('  ✗ Thiếu lớp hoặc họ tên: ' + listShort_(partial)); problems++; }
  if (dupes.length) { lines.push('  ✗ Trùng họ tên trong cùng lớp (cần thêm ký hiệu phân biệt): ' + listShort_(dupes)); problems++; }
  var inconsistent = Object.keys(variants).filter(function (f) { return Object.keys(variants[f]).length > 1; })
    .map(function (f) { return Object.keys(variants[f]).join(' / '); });
  if (inconsistent.length) { lines.push('  ✗ Tên lớp viết không thống nhất: ' + listShort_(inconsistent)); problems++; }
  if (digits.length) { lines.push('  ! Họ tên có chữ số (có thể nhầm số thứ tự): ' + listShort_(digits)); problems++; }

  // --- Câu hỏi ---
  var qs = getQSheet_();
  var qLast = qs ? qs.getLastRow() : 0;
  var raw = qLast >= 2 ? qs.getRange(2, 1, qLast - 1, Q_COLS).getValues() : [];
  var famNames = {}, entNames = {};
  readFamilies_().forEach(function (f) { famNames[fold_(f.name)] = 1; });
  readEntries_().forEach(function (e) { entNames[fold_(e.name)] = 1; });
  var firstSeen = {}, taken = {};
  raw.forEach(function (r) { var id = parseInt(r[0], 10); if (id >= 1 && id <= MAX_Q) taken[id] = 1; });
  var nextFree = function () {
    for (var k = 1; k <= MAX_Q; k++) { if (!taken[k]) { taken[k] = 1; return k; } }
    return 0;
  };
  var badStt = [], dupStt = [], hidden = 0, active = 0, noUnit = 0, noTopic = 0, unitCount = {};
  var scopeCount = {}, scopeAll = 0, scopeUnknown = {}, activeQs = [];
  var byType = { dien_tu: 0, trac_nghiem: 0, viet_lai: 0, word_entry: 0 };
  var issues = {};   // lý do -> [STT]
  var addIssue = function (reason, id) { (issues[reason] = issues[reason] || []).push(id); };
  raw.forEach(function (r, i) {
    var sentence = str_(r[1]);
    if (!sentence && !str_(r[2]) && !str_(r[3])) return;
    var rowNo = i + 2;
    var id = parseInt(r[0], 10);
    if (!(id >= 1 && id <= MAX_Q)) {
      var free = str_(r[0]) === '' ? nextFree() : 0;
      if (free) {
        id = free; qs.getRange(rowNo, 1).setValue(id); filled++;
      } else { badStt.push('dòng ' + rowNo); return; }
    } else if (firstSeen[id]) { dupStt.push('STT ' + id + ' (dòng ' + rowNo + ')'); return; }
    firstSeen[id] = 1;
    if (!sentence) { badStt.push('dòng ' + rowNo + ' thiếu câu hỏi'); return; }
    if (str_(r[8]) !== '') { hidden++; return; }
    active++;
    var q = buildQuestion_(r, id);
    if (q.typeBad) { addIssue('✗ Dạng bài không hợp lệ (chọn: Điền từ, Trắc nghiệm, Viết lại câu, Word entry): câu', id); return; }
    byType[q.type]++;
    activeQs.push(q);
    if (q.unit) unitCount[q.unit] = (unitCount[q.unit] || 0) + 1; else noUnit++;
    if (!q.topic) noTopic++;
    var prob = questionProblem_(q, famNames, entNames);
    if (prob) addIssue((prob === 'Thiếu từ gốc' || prob.indexOf('Gia đình từ') === 0 ? '! ' : '✗ ') + prob + ': câu', id);
  });
  lines.push('');
  lines.push('CÂU HỎI: ' + active + ' câu đang bật' + (hidden ? ', ' + hidden + ' câu đang ẩn' : ''));
  if (active) {
    lines.push('  Theo dạng: ' + TYPE_ORDER.map(function (t) { return TYPE_LABEL[t] + ' ' + byType[t]; }).join(' • '));
  }
  if (active) {
    var uKeys = Object.keys(unitCount).sort(natCmp_);
    if (uKeys.length) lines.push('  Theo Unit: ' + uKeys.map(function (u) { return u + ' ' + unitCount[u]; }).join(' • ') + (noUnit ? ' • (chưa gán Unit) ' + noUnit : ''));
    if (noUnit || noTopic) lines.push('  ℹ Chưa gán Unit: ' + noUnit + ' câu; chưa gán Chủ đề: ' + noTopic + ' câu (không bắt buộc; gán để học sinh dễ tìm bài ôn).');
  }
  // --- Phạm vi khối/lớp ---
  var rosterClasses = Object.keys(count), rosterKhoi = {}, rosterKeys = {};
  rosterClasses.forEach(function (c) { rosterKeys[classKey_(c)] = 1; if (khoiOf_(c)) rosterKhoi[khoiOf_(c)] = 1; });
  activeQs.forEach(function (q) {
    var toks = scopeTokens_(q.scope);
    if (!toks.length) { scopeAll++; return; }
    toks.forEach(function (t) {
      var info = scopeTokenInfo_(t);
      if (info.kind === 'all') { scopeAll++; return; }
      var label = info.kind === 'khoi' ? 'Khối ' + info.key : 'Lớp ' + t;
      scopeCount[label] = (scopeCount[label] || 0) + 1;
      var known = info.kind === 'khoi' ? !!rosterKhoi[info.key] : !!rosterKeys[info.key];
      if (!known) (scopeUnknown[label] = scopeUnknown[label] || []).push(q.id);
    });
  });
  if (active) {
    var sKeys = Object.keys(scopeCount).sort(natCmp_);
    lines.push('  Theo khối/lớp: ' + (scopeAll ? 'dùng chung (mọi lớp) ' + scopeAll : '') + (scopeAll && sKeys.length ? ' • ' : '') +
      sKeys.map(function (k) { return k + ' ' + scopeCount[k]; }).join(' • ') + (!scopeAll && !sKeys.length ? '(chưa gán)' : ''));
    Object.keys(scopeUnknown).forEach(function (k) {
      lines.push('  ! "' + k + '" không khớp khối/lớp nào trong DanhSach (sai chính tả?): câu ' + listShort_(scopeUnknown[k]));
      problems++;
    });
    rosterClasses.sort(natCmp_).forEach(function (c) {
      var cnt = activeQs.filter(function (q) { return inScope_(q.scope, c); }).length;
      if (!cnt) { lines.push('  ✗ Lớp "' + c + '" hiện KHÔNG có câu hỏi nào được giao.'); problems++; }
    });
  }
  if (!active) { lines.push('  ✗ Chưa có câu hỏi nào đang bật.'); problems++; }
  if (filled) lines.push('  ✓ Đã tự đánh số STT cho ' + filled + ' câu mới.');
  if (badStt.length) { lines.push('  ✗ STT không hợp lệ (cần số từ 1 đến ' + MAX_Q + '): ' + listShort_(badStt)); problems++; }
  if (dupStt.length) { lines.push('  ✗ STT bị trùng, câu sau sẽ KHÔNG hiện: ' + listShort_(dupStt)); problems++; }
  Object.keys(issues).forEach(function (reason) {
    lines.push('  ' + reason + ' ' + listShort_(issues[reason]));
    problems++;
  });
  if (!problems) lines.push('\n✔ Dữ liệu ổn, không phát hiện lỗi.');
  return { text: lines.join('\n'), problems: problems, filled: filled };
}

function kiemTraDuLieu() {
  alert_(checkData_().text);
}

// ============================================================================
//  10. MENU VÀ THIẾT LẬP SHEET
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Word Form 9')
    .addItem('1. Thiết lập ban đầu', 'thietLapBanDau')
    .addItem('2. Đặt mật khẩu giáo viên', 'datMatKhauGiaoVien')
    .addSeparator()
    .addItem('3. Nhập nhanh danh sách lớp (dán)', 'nhapDanhSachNhanh')
    .addItem('4. Kiểm tra dữ liệu (lớp + câu hỏi)', 'kiemTraDuLieu')
    .addItem('5. Chấm lại tất cả bài đã nộp', 'chamLaiTatCa')
    .addItem('6. Làm mới thống kê (TongQuan, ThongKeCau)', 'lamMoiTongQuan')
    .addItem('7. Xem link học sinh / link giáo viên', 'xemLink')
    .addItem('8. Nâng cấp lên 4 dạng bài (thêm, không xóa)', 'nangCapBanMoi')
    .addItem('9. Nâng cấp lên 300 câu + Unit/Chủ đề (thêm, không xóa)', 'nangCapBan6')
    .addItem('10. Nâng cấp v7: câu theo khối/lớp + 800 câu (thêm, không xóa)', 'nangCapBan7')
    .addItem('11. Nâng cấp lên 2.000 câu (thêm cột, không xóa)', 'nangCapBan8')
    .addSeparator()
    .addItem('Xóa toàn bộ kết quả (năm học mới)', 'xoaKetQua')
    .addToUi();
}

function xemLink() {
  var url = pageUrl_();
  if (!url || !/\/exec$/.test(url)) {
    alert_('Chưa có link chính thức.\n\nHãy bấm Triển khai → Tùy chọn triển khai mới → Ứng dụng web (Thực thi: Tôi; Truy cập: Bất kỳ ai), rồi mở lại menu này.');
    return;
  }
  alert_('LINK HỌC SINH (gửi cho học sinh):\n' + url +
    '\n\nLINK GIÁO VIÊN (bảng điều khiển, có mật khẩu):\n' + url + '?page=dashboard' +
    '\n\nLINK RIÊNG TỪNG LỚP / KHỐI (học sinh chỉ thấy lớp đó):\n' + url + '?lop=7/1\n' + url + '?khoi=7' +
    '\n(Bảng link đủ các lớp: trang giáo viên, thẻ "Link lớp")');
}

function alert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function colLetter_(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function styleHeader_(range) {
  range.setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

/** Tiêu đề, định dạng, danh sách chọn "Dạng bài" và ghi chú cho sheet CauHoi (chạy lại nhiều lần vẫn an toàn). */
function setupQuestionSheet_(qs) {
  ensureCols_(qs, Q_COLS);
  qs.getRange(1, 1, 1, Q_COLS).setValues([['STT', 'Câu hỏi (Điền từ / Word entry: gõ ____ vào chỗ trống)', 'Từ gốc (Điền từ) hoặc Gợi ý bắt đầu (Viết lại câu)',
    'Đáp án (bắt buộc; trắc nghiệm: A, B, C hoặc D)', 'Đáp án khác chấp nhận (ngăn cách bằng |)',
    'Giải thích ngữ pháp', 'Dịch nghĩa', 'Gia đình từ (Điền từ; tên ở sheet GiaDinhTu)', 'Ẩn câu này (gõ x)',
    'Dạng bài', 'Lựa chọn A', 'Lựa chọn B', 'Lựa chọn C', 'Lựa chọn D', 'Mục từ điển (Word entry; tên ở sheet TuDien)',
    'Unit (để học sinh lọc bài)', 'Chủ đề / chuyên đề (VD: Danh từ, Câu bị động)', 'Áp dụng cho khối/lớp (VD: 7 hoặc 6|7 hoặc 9/1; để trống = mọi lớp)']]);
  styleHeader_(qs.getRange(1, 1, 1, Q_COLS));
  qs.setFrozenRows(1);
  qs.setRowHeight(1, 60);
  qs.setColumnWidths(1, 1, 50); qs.setColumnWidths(2, 1, 420); qs.setColumnWidths(3, 1, 130);
  qs.setColumnWidths(4, 1, 160); qs.setColumnWidths(5, 1, 150); qs.setColumnWidths(6, 1, 380);
  qs.setColumnWidths(7, 1, 380); qs.setColumnWidths(8, 1, 150); qs.setColumnWidths(9, 1, 80);
  qs.setColumnWidths(10, 1, 110); qs.setColumnWidths(11, 4, 170); qs.setColumnWidths(15, 1, 170);
  qs.setColumnWidths(16, 1, 220); qs.setColumnWidths(17, 1, 200); qs.setColumnWidths(18, 1, 190);
  qs.getRange(2, 2, 1000, Q_COLS - 1).setNumberFormat('@').setWrap(true).setVerticalAlignment('top');
  try {
    qs.getRange(2, 10, 1000, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(TYPE_ORDER.map(function (t) { return TYPE_LABEL[t]; }), true)
      .setAllowInvalid(false).build());
  } catch (e) {}
  try {
    qs.getRange(2, 16, 1000, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(UNIT_SUGGEST, true).setAllowInvalid(true).build());
  } catch (e) {}
  try {
    qs.getRange(1, 2).setNote('Điền từ và Word entry: gõ 4 dấu gạch dưới trở lên (____) vào đúng chỗ học sinh cần điền, mỗi câu chỉ 1 chỗ trống.\nTrắc nghiệm: gõ câu hỏi (có thể có ____). Viết lại câu: gõ câu gốc.');
    qs.getRange(1, 3).setNote('Điền từ: từ gốc in hoa cho trong ngoặc (VD: ATTRACT).\nViết lại câu: gợi ý cho học sinh (VD: Begin with "He has"). Các dạng khác: để trống.');
    qs.getRange(1, 4).setNote('Bắt buộc có đáp án: câu chưa có đáp án sẽ KHÔNG hiện cho học sinh. Sửa đáp án xong thì chọn menu "5. Chấm lại tất cả bài đã nộp".\nTrắc nghiệm: chỉ gõ chữ cái A, B, C hoặc D. Viết lại câu: gõ cả câu đúng.');
    qs.getRange(1, 5).setNote('Các đáp án khác cũng được tính đúng, ví dụ: Prioritising|prioritizing. Viết lại câu: ghi thêm các cách viết đúng khác.');
    qs.getRange(1, 9).setNote('Gõ x để tạm ẩn câu hỏi (soạn nháp). Xóa x để hiện lại.');
    qs.getRange(1, 10).setNote('Chọn: Điền từ, Trắc nghiệm, Viết lại câu hoặc Word entry. Để trống = Điền từ.');
    qs.getRange(1, 11).setNote('Chỉ dùng cho Trắc nghiệm: nội dung 4 lựa chọn A, B, C, D (ít nhất 2).');
    qs.getRange(1, 16).setNote('Unit của sách (VD: Unit 3: Teen Stress and Pressure) hoặc "Ôn tổng hợp". Học sinh lọc câu theo Unit. Chọn trong danh sách hoặc tự gõ tên mới; các câu cùng Unit cần gõ GIỐNG NHAU. Để trống = "Chưa phân Unit".');
    qs.getRange(1, 17).setNote('Chủ đề / chuyên đề ngữ pháp - từ vựng (VD: Danh từ, Trạng từ, Câu bị động, Câu điều kiện, Phát âm...). Học sinh lọc theo chủ đề. Để trống nếu chưa phân loại.');
    qs.getRange(1, 18).setNote('Khối/lớp được làm câu này. Gõ số khối (6, 7, 8, 9) hoặc tên lớp đúng như ở DanhSach (VD 7/1 hoặc 9A1). Nhiều mục cách nhau dấu | (VD: 6|7 hoặc 9/1|9/2). Câu dùng chung nhiều khối: gõ cả các khối. Để TRỐNG = mọi lớp đều thấy.');
    qs.getRange(1, 15).setNote('Chỉ dùng cho Word entry: gõ đúng tên mục từ ở cột A của sheet TuDien. Hai câu Word entry có thể dùng chung một mục từ.');
  } catch (e) {}
}

/** Tiêu đề và định dạng cho sheet TuDien (mục từ điển cho dạng Word entry). */
function setupEntrySheet_(en) {
  en.getRange(1, 1, 1, 5).setValues([['Mục từ (headword)', 'Phiên âm IPA', 'Từ loại', 'Nghĩa (mỗi nghĩa 1 dòng, Alt+Enter)',
    'Câu ví dụ (mỗi câu 1 dòng, Alt+Enter; cụm cần học đặt giữa hai dấu ** , VD: **make a decision**)']]);
  styleHeader_(en.getRange(1, 1, 1, 5));
  en.setFrozenRows(1);
  en.setRowHeight(1, 60);
  en.setColumnWidths(1, 1, 150); en.setColumnWidths(2, 1, 120); en.setColumnWidths(3, 1, 90);
  en.setColumnWidths(4, 1, 330); en.setColumnWidths(5, 1, 560);
  en.getRange(2, 1, 300, 5).setNumberFormat('@').setWrap(true).setVerticalAlignment('top');
}

/**
 * Menu "8. Nâng cấp lên 4 dạng bài": dành cho Sheet đã thiết lập bằng bản cũ.
 * Chỉ THÊM (cột Dạng bài, lựa chọn, mục từ điển, sheet TuDien, câu mẫu mới); không sửa, không xóa câu hay kết quả đã có.
 */
function nangCapBanMoi() {
  var ss = SpreadsheetApp.getActive();
  var qs = ss.getSheetByName(SH_QUESTIONS);
  if (!qs) { alert_('Chưa thiết lập. Hãy chạy menu 1. Thiết lập ban đầu trước.'); return; }
  var res = withLock_(function () {
    var report = [];
    // 1) sheet TuDien
    var en = getOrCreate_(ss, SH_ENTRY);
    var addedEntries = 0;
    var haveEntry = {};
    readEntries_().forEach(function (e) { haveEntry[fold_(e.name)] = 1; });
    var newEntries = SAMPLE_ENTRIES.filter(function (e) { return !haveEntry[fold_(e[0])]; });
    if (newEntries.length) {
      var er = Math.max(en.getLastRow(), 1) + 1;
      en.getRange(er, 1, newEntries.length, 5).setNumberFormat('@').setValues(newEntries);
      addedEntries = newEntries.length;
    }
    setupEntrySheet_(en);
    // 2) cột mới của CauHoi (chỉ ghi tiêu đề/định dạng cột J..O, không đụng dữ liệu A..I)
    setupQuestionSheet_(qs);
    // 3) câu mẫu mới: thêm vào các STT còn trống, bỏ qua câu đã có sẵn
    var lr = qs.getLastRow();
    var data = lr >= 2 ? qs.getRange(2, 1, lr - 1, Q_COLS).getValues() : [];
    var taken = {}, haveText = {};
    data.forEach(function (r) {
      var id = parseInt(r[0], 10);
      if (id >= 1 && id <= MAX_Q) taken[id] = 1;
      if (str_(r[1])) haveText[fold_(r[1])] = 1;
    });
    var toAdd = [];
    SAMPLE_QUESTIONS_NEW.forEach(function (row) {
      if (haveText[fold_(row[1])]) return;
      var id = lowestFree_(taken);
      if (!id) return;
      taken[id] = 1;
      var copy = padRow_(row); copy[0] = id;
      toAdd.push(copy);
    });
    if (toAdd.length) {
      var start = Math.max(lr, 1) + 1;
      qs.getRange(start, 2, toAdd.length, Q_COLS - 1).setNumberFormat('@');
      qs.getRange(start, 1, toAdd.length, Q_COLS).setValues(toAdd);
    }
    // 4) câu cũ để trống cột Dạng bài vẫn là Điền từ (không cần sửa)
    return { entries: addedEntries, questions: toAdd.length };
  });
  refreshSummary_();
  alert_('Đã nâng cấp lên 4 dạng bài.\n- Sheet TuDien: ' + (res.entries ? 'đã thêm ' + res.entries + ' mục từ mẫu (dữ liệu có sẵn giữ nguyên)' : 'giữ nguyên dữ liệu có sẵn') +
    '\n- Câu hỏi mẫu mới thêm: ' + res.questions + ' (trắc nghiệm, viết lại câu, word entry)\n- Các câu và kết quả cũ không bị thay đổi.\nHãy chạy menu 4. Kiểm tra dữ liệu.');
}

/**
 * Menu "9. Nâng cấp lên 300 câu + Unit/Chủ đề": dành cho Sheet đang chạy bản cũ.
 * Chỉ THÊM: cột Unit, Chủ đề (P, Q) của CauHoi và các cột Câu 101-300 của TienDo.
 * Không sửa/xóa câu hỏi, kết quả hay danh sách lớp. Chỉ điền Unit/Chủ đề cho các câu MẪU còn để trống ô này.
 */
function nangCapBan6() {
  var ss = SpreadsheetApp.getActive();
  var qs = ss.getSheetByName(SH_QUESTIONS);
  var pr = ss.getSheetByName(SH_PROGRESS);
  if (!qs || !pr) { alert_('Chưa thiết lập. Hãy chạy menu 1. Thiết lập ban đầu trước.'); return; }
  var res = withLock_(function () {
    // 1) CauHoi: thêm cột P, Q (tiêu đề, định dạng, danh sách chọn Unit)
    setupQuestionSheet_(qs);
    // 2) TienDo: thêm các cột Câu 101..300 (giữ nguyên cột JSON ở vị trí cũ)
    ensureCols_(pr, PROG_COLS);
    var head = [];
    for (var id = Q_BLOCK1 + 1; id <= MAX_Q; id++) head.push('Câu ' + id);
    pr.getRange(1, COL_TEXTS + 1, 1, head.length).setValues([head]);
    styleHeader_(pr.getRange(1, COL_TEXTS + 1, 1, head.length));
    formatProgressGrid_(pr);
    // 3) Unit/Chủ đề gợi ý cho câu mẫu (chỉ ô đang trống)
    var meta = {};
    sampleRows_().forEach(function (r) { meta[fold_(r[1])] = [r[15], r[16]]; });
    var lr = qs.getLastRow(), filled = 0;
    if (lr >= 2) {
      var data = qs.getRange(2, 1, lr - 1, Q_COLS).getValues();
      var out = [];
      data.forEach(function (r) {
        var m = meta[fold_(r[1])];
        var u = str_(r[15]), t = str_(r[16]);
        if (m && !u && !t) { u = m[0]; t = m[1]; filled++; }
        out.push([r[15] === '' || r[15] == null ? u : r[15], r[16] === '' || r[16] == null ? t : r[16]]);
      });
      if (filled) {
        qs.getRange(2, 16, out.length, 2).setNumberFormat('@');
        qs.getRange(2, 16, out.length, 2).setValues(out);
      }
    }
    return { filled: filled };
  });
  refreshSummary_();
  alert_('Đã nâng cấp xong.\n- Số câu tối đa: ' + MAX_Q + ' (trước đây 100).\n- CauHoi có thêm cột P "Unit" và Q "Chủ đề"' +
    (res.filled ? '; đã gợi ý sẵn Unit/Chủ đề cho ' + res.filled + ' câu mẫu (bạn có thể sửa)' : '') +
    '.\n- TienDo đã thêm các cột Câu 101-' + MAX_Q + '.\n- Câu hỏi, kết quả và danh sách lớp cũ không bị thay đổi.\nHãy chạy menu 4. Kiểm tra dữ liệu.');
}

/**
 * Menu "10. Nâng cấp v7": dành cho Sheet đang chạy v6.
 * Chỉ THÊM: cột R "Áp dụng cho" của CauHoi và các cột Câu 101-800 của TienDo.
 * Không sửa/xóa câu hỏi, kết quả hay danh sách lớp. Câu cũ có cột R trống = mọi lớp vẫn thấy như trước.
 */
function nangCapBan7() {
  var ss = SpreadsheetApp.getActive();
  var qs = ss.getSheetByName(SH_QUESTIONS);
  var pr = ss.getSheetByName(SH_PROGRESS);
  if (!qs || !pr) { alert_('Chưa thiết lập. Hãy chạy menu 1. Thiết lập ban đầu trước.'); return; }
  withLock_(function () {
    setupQuestionSheet_(qs);
    ensureCols_(pr, PROG_COLS);
    var head = [];
    for (var id = Q_BLOCK1 + 1; id <= MAX_Q; id++) head.push('Câu ' + id);
    pr.getRange(1, COL_TEXTS + 1, 1, head.length).setValues([head]);
    styleHeader_(pr.getRange(1, COL_TEXTS + 1, 1, head.length));
    formatProgressGrid_(pr);
  });
  refreshSummary_();
  alert_('Đã nâng cấp lên v7.\n- Số câu tối đa: ' + MAX_Q + '.\n- CauHoi có thêm cột R "Áp dụng cho (khối/lớp)": gõ 6, 7, 8, 9 (khối) hoặc tên lớp (VD 7/1); nhiều mục cách nhau dấu |. Để trống = mọi lớp đều thấy (như trước đây).\n- TienDo đã có đủ cột Câu 1-' + MAX_Q + '.\n- Câu hỏi, kết quả và danh sách lớp cũ không bị thay đổi.\nHãy chạy menu 4. Kiểm tra dữ liệu.');
}

/**
 * Menu "11. Nâng cấp lên 2.000 câu": chỉ THÊM cột "Câu 801-2000" vào TienDo (kèm tiêu đề, định dạng).
 * Không sửa/xóa câu hỏi, kết quả hay danh sách lớp. Chạy được nhiều lần.
 */
function nangCapBan8() {
  var ss = SpreadsheetApp.getActive();
  var qs = ss.getSheetByName(SH_QUESTIONS);
  var pr = ss.getSheetByName(SH_PROGRESS);
  if (!qs || !pr) { alert_('Chưa thiết lập. Hãy chạy menu 1. Thiết lập ban đầu trước.'); return; }
  withLock_(function () {
    ensureCols_(pr, PROG_COLS);
    var head = [];
    for (var id = Q_BLOCK1 + 1; id <= MAX_Q; id++) head.push('Câu ' + id);
    pr.getRange(1, COL_TEXTS + 1, 1, head.length).setValues([head]);
    styleHeader_(pr.getRange(1, COL_TEXTS + 1, 1, head.length));
    formatProgressGrid_(pr);
  });
  refreshSummary_();
  alert_('Đã nâng cấp xong.\n- Số câu tối đa: ' + MAX_Q + '.\n- TienDo đã có thêm các cột Câu 801-' + MAX_Q + ' (các cột chưa dùng được ẩn tự động).\n- Câu hỏi, kết quả và danh sách lớp cũ không bị thay đổi.\nHãy chạy menu 4. Kiểm tra dữ liệu.');
}

function thietLapBanDau() {
  var ss = SpreadsheetApp.getActive();
  getSecret_();

  // --- DanhSach ---
  var ro = getOrCreate_(ss, SH_ROSTER);
  ro.getRange('A:C').setNumberFormat('@');
  if (ro.getLastRow() < 1) {
    ro.getRange(1, 1, 1, 4).setValues([['Lớp', 'Họ và tên', 'Mã học sinh (không bắt buộc)', 'Ghi chú']]);
    ro.getRange(2, 1, 1, 4).setValues([['9A1', 'Nguyễn Văn An', '', 'VÍ DỤ - hãy xóa dòng này']]);
  }
  styleHeader_(ro.getRange(1, 1, 1, 4));
  ro.setFrozenRows(1);
  ro.setColumnWidths(1, 1, 90); ro.setColumnWidths(2, 1, 240);
  ro.setColumnWidths(3, 1, 190); ro.setColumnWidths(4, 1, 240);

  // --- GiaDinhTu ---
  var fa = getOrCreate_(ss, SH_FAMILY);
  if (fa.getLastRow() < 2) {
    fa.getRange(1, 1, 1, 2).setValues([['Từ gốc (tên nhóm)', 'Nội dung gia đình từ (dùng dấu → để ngăn cách các dạng từ)']]);
    fa.getRange(2, 1, SAMPLE_FAMILIES.length, 2).setValues(SAMPLE_FAMILIES);
  }
  styleHeader_(fa.getRange(1, 1, 1, 2));
  fa.setFrozenRows(1);
  fa.setColumnWidths(1, 1, 170); fa.setColumnWidths(2, 1, 700);
  fa.getRange(2, 2, 500, 1).setWrap(true);

  // --- TuDien + CauHoi ---
  var en = getOrCreate_(ss, SH_ENTRY);
  if (en.getLastRow() < 2) en.getRange(2, 1, SAMPLE_ENTRIES.length, 5).setNumberFormat('@').setValues(SAMPLE_ENTRIES);
  setupEntrySheet_(en);
  var qs = getOrCreate_(ss, SH_QUESTIONS);
  if (qs.getLastRow() < 2) {
    var rows = sampleRows_();
    qs.getRange(2, 2, rows.length, Q_COLS - 1).setNumberFormat('@');
    qs.getRange(2, 1, rows.length, Q_COLS).setValues(rows);
  }
  setupQuestionSheet_(qs);

  // --- TienDo ---
  var pr = ensureCols_(getOrCreate_(ss, SH_PROGRESS), PROG_COLS);
  var head = ['Lớp', 'Họ và tên', 'Đã làm', 'Đúng', 'Sai', 'Chờ chấm', 'Tiến độ', '% đúng', 'Điểm /10', 'Tổng lượt', 'Lần cuối'];
  while (head.length < PROG_COLS) head.push('');
  for (var q = 1; q <= MAX_Q; q++) head[qCol_(q) - 1] = 'Câu ' + q;
  head[COL_TEXTS - 1] = 'Bài làm (JSON - đừng sửa)';
  pr.getRange(1, 1, 1, PROG_COLS).setValues([head]);
  styleHeader_(pr.getRange(1, 1, 1, PROG_COLS));
  pr.getRange('A:B').setNumberFormat('@');
  pr.getRange('G:H').setNumberFormat('0%');
  pr.getRange('I:I').setNumberFormat('0.0');
  pr.getRange('K:K').setNumberFormat('dd/MM/yyyy HH:mm');
  pr.setFrozenRows(1); pr.setFrozenColumns(2);
  pr.setColumnWidths(1, 1, 70); pr.setColumnWidths(2, 1, 200);
  pr.setColumnWidths(3, 8, 66); pr.setColumnWidths(11, 1, 130);
  formatProgressGrid_(pr);
  try { pr.getRange(1, 3).setNote('Đ = đúng • S = sai • X = đã bấm Xem đáp án'); } catch (e) {}
  try { pr.protect().setWarningOnly(true); } catch (e) {}

  // --- NhatKy ---
  var lg = getOrCreate_(ss, SH_LOG);
  lg.getRange(1, 1, 1, 9).setValues([['Thời gian', 'Lớp', 'Họ và tên', 'Hình thức', 'Câu', 'Từ gốc', 'Bài làm', 'Kết quả', 'Mã gói']]);
  styleHeader_(lg.getRange(1, 1, 1, 9));
  lg.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  lg.getRange('B:C').setNumberFormat('@');
  lg.setFrozenRows(1);
  lg.setColumnWidths(1, 1, 150); lg.setColumnWidths(3, 1, 200); lg.setColumnWidths(4, 1, 140);
  try { lg.protect().setWarningOnly(true); } catch (e) {}

  lamMoiTongQuan();

  if (!PropertiesService.getScriptProperties().getProperty('TEACHER_KEY')) datMatKhauGiaoVien();

  alert_('Đã thiết lập xong. Việc tiếp theo:\n1) Nhập danh sách học sinh: menu 3. Nhập nhanh danh sách lớp (hoặc gõ thẳng vào sheet DanhSach, nhớ xóa dòng ví dụ).\n2) Câu hỏi: đã có sẵn 40 câu của file gốc. Thêm/sửa/xóa ngay trên trang giáo viên (hoặc sheet CauHoi).\n3) Menu 4. Kiểm tra dữ liệu.\n4) Triển khai Web App (xem HUONG_DAN.md).');
}

/** Độ rộng, căn giữa, tô màu Đ/S/X cho các cột câu hỏi của TienDo (cả hai khối: Câu 1-100 và Câu 101-300). */
function formatProgressGrid_(pr) {
  pr.setColumnWidths(COL_Q0, Q_BLOCK1, 46);
  pr.setColumnWidths(COL_TEXTS + 1, MAX_Q - Q_BLOCK1, 46);
  pr.hideColumns(COL_TEXTS);
  var g1 = pr.getRange(2, COL_Q0, 998, Q_BLOCK1);
  var g2 = pr.getRange(2, COL_TEXTS + 1, 998, MAX_Q - Q_BLOCK1);
  g1.setHorizontalAlignment('center');
  g2.setHorizontalAlignment('center');
  var rg = [g1, g2];
  var rule = function (txt, color) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(txt).setBackground(color).setRanges(rg).build();
  };
  pr.setConditionalFormatRules([
    rule(ST_OK, '#b7e1cd'), rule(ST_BAD, '#f4b6b6'), rule(ST_REV, '#fde68a'),
    rule(ST_PEND, '#e2e8f0'), rule(ST_SELF_OK, '#d1fae5'), rule(ST_SELF_BAD, '#fde2e2')
  ]);
}

function datMatKhauGiaoVien() {
  var props = PropertiesService.getScriptProperties();
  var key = '';
  try {
    var ui = SpreadsheetApp.getUi();
    var r = ui.prompt('Mật khẩu giáo viên', 'Nhập mật khẩu để mở trang bảng điều khiển (ít nhất 6 ký tự):', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    key = String(r.getResponseText() || '').trim();
    if (key.length < 6) { ui.alert('Mật khẩu quá ngắn (cần ít nhất 6 ký tự). Hãy chạy lại bước này.'); return; }
  } catch (e) {
    key = Utilities.getUuid().slice(0, 8);
    Logger.log('Mật khẩu giáo viên được tạo tự động: ' + key);
  }
  props.setProperty('TEACHER_KEY', key);
  alert_('Đã lưu mật khẩu giáo viên.');
}

function chamLaiTatCa() {
  var r = regradeAll_();
  lamMoiTongQuan();
  alert_('Đã chấm lại xong.\nSố học sinh: ' + r.students + '\nSố ô kết quả thay đổi: ' + r.cells);
}

function lamMoiTongQuan() {
  var ss = SpreadsheetApp.getActive();
  var questions = readQuestions_(false);
  var maxId = 0;
  questions.forEach(function (q) { if (q.id > maxId) maxId = q.id; });

  // --- Ẩn/hiện cột câu hỏi trong TienDo cho vừa số câu hiện có ---
  var pr = ss.getSheetByName(SH_PROGRESS);
  if (pr) {
    try {
      ensureCols_(pr, PROG_COLS);
      var n1 = Math.min(maxId, Q_BLOCK1), n2 = Math.max(0, maxId - Q_BLOCK1);
      if (n1 > 0) pr.showColumns(COL_Q0, n1);
      if (n1 < Q_BLOCK1) pr.hideColumns(COL_Q0 + n1, Q_BLOCK1 - n1);
      if (n2 > 0) pr.showColumns(COL_TEXTS + 1, n2);
      if (n2 < MAX_Q - Q_BLOCK1) pr.hideColumns(COL_TEXTS + 1 + n2, MAX_Q - Q_BLOCK1 - n2);
      pr.hideColumns(COL_TEXTS);
    } catch (e) {}
  }

  // --- TongQuan: thống kê theo lớp ---
  var ov = getOrCreate_(ss, SH_OVERVIEW);
  ov.clear();
  ov.getRange(1, 1, 1, 7).setValues([['Lớp', 'Sĩ số', 'Đã đăng nhập', 'Đã làm ≥ 1 câu', 'Hoàn thành hết', 'Điểm TB (thang 10)', '% đúng TB']]);
  styleHeader_(ov.getRange(1, 1, 1, 7));
  var classes = listClasses_().classes;
  if (classes.length) {
    ov.getRange(2, 1, classes.length, 1).setNumberFormat('@').setValues(classes.map(function (c) { return [c]; }));
    var f = [];
    for (var i = 0; i < classes.length; i++) {
      var r = i + 2;
      var A = '(TienDo!$A$2:$A=$A' + r + ')';
      var started = 'ISNUMBER(TienDo!$C$2:$C)*(TienDo!$C$2:$C>0)';
      f.push([
        '=SUMPRODUCT(--(DanhSach!$A$2:$A=$A' + r + '))',
        '=SUMPRODUCT(--' + A + ')',
        '=SUMPRODUCT(' + A + '*' + started + ')',
        '=SUMPRODUCT(' + A + '*ISNUMBER(TienDo!$G$2:$G)*(TienDo!$G$2:$G>=1))',
        '=IFERROR(SUMPRODUCT(' + A + '*' + started + '*TienDo!$I$2:$I)/D' + r + ',0)',
        '=IFERROR(SUMPRODUCT(' + A + '*' + started + '*TienDo!$H$2:$H)/D' + r + ',0)'
      ]);
    }
    ov.getRange(2, 2, classes.length, 6).setValues(f);
    ov.getRange(2, 6, classes.length, 1).setNumberFormat('0.0');
    ov.getRange(2, 7, classes.length, 1).setNumberFormat('0%');
  }
  ov.setFrozenRows(1);
  ov.setColumnWidths(1, 7, 130);

  // --- ThongKeCau ---
  var qs = getOrCreate_(ss, SH_QSTATS);
  var oldF = qs.getFilter();
  if (oldF) oldF.remove();
  qs.clear();
  qs.getRange(1, 1, 1, 8).setValues([['Câu', 'Từ gốc / dạng bài', 'Đáp án', 'Số HS đã làm', 'Số HS đúng', 'Số HS sai / xem đáp án', 'Chờ chấm', '% sai']]);
  styleHeader_(qs.getRange(1, 1, 1, 8));
  if (questions.length) {
    var rows = questions.map(function (q, i) {
      var L = colLetter_(qCol_(q.id));
      var rng = 'TienDo!' + L + '$2:' + L;
      var rr = i + 2;
      var cnt = function (s) { return 'COUNTIF(' + rng + ',"' + s + '")'; };
      return [
        q.id, qLabel_(q), q.answer || '(chưa có đáp án)',
        '=COUNTA(' + rng + ')',
        '=' + cnt(ST_OK) + '+' + cnt(ST_SELF_OK),
        '=' + cnt(ST_BAD) + '+' + cnt(ST_REV) + '+' + cnt(ST_SELF_BAD),
        '=' + cnt(ST_PEND),
        '=IF((E' + rr + '+F' + rr + ')=0,0,F' + rr + '/(E' + rr + '+F' + rr + '))'
      ];
    });
    qs.getRange(2, 1, rows.length, 8).setValues(rows);
    qs.getRange(2, 8, rows.length, 1).setNumberFormat('0%');
    qs.getRange(1, 1, rows.length + 1, 8).createFilter();
  }
  qs.setFrozenRows(1);
  qs.setColumnWidths(1, 8, 120);
}

function xoaKetQua() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('XÓA TOÀN BỘ KẾT QUẢ',
    'Thao tác này xóa mọi dòng trong sheet TienDo và NhatKy (KHÔNG xóa DanhSach, CauHoi) và KHÔNG khôi phục được.\nGõ XOA để xác nhận:',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK || String(r.getResponseText()).trim() !== 'XOA') {
    ui.alert('Đã hủy, không có gì bị xóa.');
    return;
  }
  var ss = SpreadsheetApp.getActive();
  [SH_PROGRESS, SH_LOG].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  });
  ui.alert('Đã xóa kết quả. Danh sách lớp và câu hỏi được giữ nguyên.');
}
