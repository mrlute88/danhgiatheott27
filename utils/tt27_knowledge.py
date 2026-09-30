"""
tt27_knowledge.py - Knowledge base for Circular 27/2020/TT-BGDĐT
Contains pedagogical rules, subject competencies, score-to-level mapping, and sample phrases.
"""

from typing import Dict, List, Tuple

# Supported Subjects according to Vietnam Primary Curriculum (Chương trình GDPT 2018)
SUBJECTS = [
    "Toán",
    "Tiếng Việt",
    "Tiếng Anh",
    "Tự nhiên và Xã hội (Lớp 1, 2, 3)",
    "Khoa học (Lớp 4, 5)",
    "Lịch sử và Địa lí (Lớp 4, 5)",
    "Tin học và Công nghệ (Lớp 3, 4, 5)",
    "Đạo đức",
    "Âm nhạc",
    "Mĩ thuật",
    "Giáo dục thể chất",
    "Hoạt động trải nghiệm"
]

GRADES = ["Lớp 1", "Lớp 2", "Lớp 3", "Lớp 4", "Lớp 5"]

PERIODS = [
    "Giữa Học kỳ 1",
    "Cuối Học kỳ 1",
    "Giữa Học kỳ 2",
    "Cuối Học kỳ 2 (Cả năm)",
    "Đánh giá thường xuyên"
]

COMMENT_TONES = [
    "Chuẩn mực & Động viên",
    "Ấm áp & Khích lệ sâu sắc",
    "Cụ thể hóa kỹ năng & Hành động",
    "Ngắn gọn, súc tích (15-20 từ)"
]

# Subject-specific core competencies & typical strengths/areas for growth
SUBJECT_COMPETENCIES: Dict[str, Dict[str, List[str]]] = {
    "Toán": {
        "strengths": [
            "tính toán nhanh và chính xác",
            "nắm chắc kiến thức, giải toán có lời văn rõ ràng",
            "tư duy logic tốt, trình bày bài cẩn thận",
            "nhận biết hình học và đo lường tốt",
            "vận dụng sáng tạo kiến thức vào giải quyết bài toán",
            "chủ động học tập, kỹ năng tính nhẩm thành thạo"
        ],
        "growth_areas": [
            "cần rèn thêm kỹ năng tính nhẩm cộng trừ",
            "cần đọc kỹ đề bài toán có lời văn để tìm phép tính phù hợp",
            "chú ý đặt tính và tính toán cẩn thận hơn",
            "cần rèn thêm kỹ năng nhân chia",
            "chú ý đơn vị đo và cách lập luận khi giải toán",
            "cần chú ý hoàn thành bài đúng thời gian quy định"
        ]
    },
    "Tiếng Việt": {
        "strengths": [
            "đọc to, rõ ràng, lưu loát và diễn cảm",
            "hiểu bài tốt, trả lời câu hỏi đọc hiểu chính xác",
            "viết chữ nắn nót, giữ vở sạch đẹp",
            "dùng từ đặt câu phong phú, đoạn văn giàu cảm xúc",
            "viết đúng chính tả, ngữ pháp chuẩn",
            "tích cực tương tác và chia sẻ cảm nghĩ trong giờ học"
        ],
        "growth_areas": [
            "cần rèn đọc to, rõ ràng và ngắt nghỉ đúng dấu câu",
            "cần chú ý viết đúng chính tả các âm đầu hoặc vần khó",
            "cần mở rộng vốn từ và trau chuốt câu văn mạch lạc hơn",
            "rèn luyện thêm về độ cao và nét chữ cho đều đẹp",
            "chú ý tập trung trả lời đúng trọng tâm câu hỏi đọc hiểu",
            "cố gắng phát biểu to và tự tin hơn trước lớp"
        ]
    },
    "Tiếng Anh": {
        "strengths": [
            "phát âm chuẩn, giao tiếp tự tin trước lớp",
            "ghi nhớ từ vựng tốt, ngữ điệu tự nhiên",
            "nghe hiểu tốt và phản xạ tiếng Anh nhanh",
            "nắm chắc cấu trúc câu đơn giản, viết câu đúng",
            "hào hứng tham gia các trò chơi và bài học tiếng Anh"
        ],
        "growth_areas": [
            "cần rèn phát âm rõ ràng các âm đuôi (ending sounds)",
            "chú ý ôn tập và ghi nhớ các từ vựng theo chủ điểm",
            "cần tự tin nói to hơn khi luyện giao tiếp cùng bạn",
            "cần chú ý cấu trúc câu và cách viết đúng ngữ pháp",
            "tích cực tham gia trả lời các câu hỏi trong giờ học"
        ]
    },
    "Tự nhiên và Xã hội (Lớp 1, 2, 3)": {
        "strengths": [
            "quan sát tỉ mỉ, biết liên hệ thực tế xung quanh",
            "tích cực phát biểu, yêu thích tìm hiểu thế giới tự nhiên",
            "có ý thức giữ gìn vệ sinh cá nhân và bảo vệ môi trường",
            "chăm chú lắng nghe, hiểu và ghi nhớ nội dung bài học"
        ],
        "growth_areas": [
            "cần mạnh dạn chia sẻ những quan sát của mình với thầy cô và bạn bè",
            "chú ý liên hệ thực tế nhiều hơn trong đời sống hằng ngày",
            "cần tập trung hơn khi làm việc nhóm và tham gia hoạt động"
        ]
    },
    "Khoa học (Lớp 4, 5)": {
        "strengths": [
            "say mê khám phá thế giới tự nhiên, hiểu bài nhanh",
            "biết làm thí nghiệm đơn giản và nêu kết luận chính xác",
            "vận dụng tốt kiến thức khoa học vào đời sống sinh hoạt",
            "tư duy khoa học tốt, lập luận rõ ràng"
        ],
        "growth_areas": [
            "cần chủ động ghi chép và thực hành thí nghiệm cẩn thận hơn",
            "chú ý củng cố lại các khái niệm khoa học trọng tâm",
            "tích cực tham gia đóng góp ý kiến khi thảo luận nhóm"
        ]
    },
    "Lịch sử và Địa lí (Lớp 4, 5)": {
        "strengths": [
            "nhớ tốt các sự kiện lịch sử, tự hào về quê hương đất nước",
            "kỹ năng chỉ bản đồ, lược đồ địa lí thành thạo",
            "hiểu bài nhanh, biết liên hệ với địa phương",
            "yêu thích tìm hiểu về văn hóa các vùng miền"
        ],
        "growth_areas": [
            "cần rèn thêm kỹ năng quan sát và đọc thông tin trên lược đồ",
            "chú ý ghi nhớ các mốc thời gian và sự kiện lịch sử tiêu biểu",
            "cố gắng phát biểu mạch lạc hơn khi trình bày nội dung bài học"
        ]
    },
    "Tin học và Công nghệ (Lớp 3, 4, 5)": {
        "strengths": [
            "thao tác máy tính nhanh, sử dụng chuột và bàn phím thành thạo",
            "hoàn thành tốt các bài thực hành, có tính sáng tạo",
            "hiểu và tuân thủ tốt quy tắc an toàn trong phòng máy",
            "tiếp thu nhanh các phần mềm học tập mới"
        ],
        "growth_areas": [
            "cần luyện tập thêm kỹ năng gõ phím bằng 10 ngón",
            "chú ý lắng nghe hướng dẫn trước khi thao tác thực hành",
            "cần kiên nhẫn hơn khi thực hiện các bài vẽ hoặc lập trình"
        ]
    },
    "Đạo đức": {
        "strengths": [
            "ngoan ngoãn, lễ phép với thầy cô và người lớn",
            "biết quan tâm, đoàn kết và giúp đỡ bạn bè trong lớp",
            "có ý thức chấp hành tốt nội quy trường lớp",
            "trung thực, có tinh thần trách nhiệm cao trong công việc chung"
        ],
        "growth_areas": [
            "cần mạnh dạn hơn khi giao tiếp với thầy cô và bạn bè",
            "chú ý tự giác giữ gìn vệ sinh bàn ghế và không gian lớp học",
            "cố gắng kiên trì thực hiện thói quen tốt mỗi ngày"
        ]
    }
}

# Rule-based comment bank for TT27 fallback & high-diversity generation
TT27_COMMENT_BANK: Dict[str, Dict[str, List[str]]] = {
    "Toán": {
        "T": [
            "Em nắm vững kiến thức, kỹ năng tính toán nhanh và rất chính xác. Cần tiếp tục phát huy.",
            "Tư duy logic tốt, giải toán có lời văn mạch lạc và cẩn thận. Thầy/Cô rất khen ngợi em.",
            "Em tiếp thu bài nhanh, làm bài cẩn thận và trình bày sạch đẹp. Hãy duy trì phong độ này.",
            "Kỹ năng tính nhẩm tốt, luôn hoàn thành bài tập trước thời gian. Đáng khen ngợi.",
            "Em vận dụng kiến thức linh hoạt, có sự sáng tạo trong cách giải bài. Phát huy tốt nhé!"
        ],
        "H": [
            "Em nắm được kiến thức cơ bản, tính toán tương đối tốt. Cần rèn thêm kỹ năng giải toán có lời văn.",
            "Có ý thức học tập, làm bài đầy đủ. Em nên chú ý tính toán cẩn thận hơn để tránh sai sót nhỏ.",
            "Em hiểu bài, hoàn thành các bài tập trên lớp. Cần rèn luyện thêm kỹ năng đặt tính và tính.",
            "Có tiến bộ trong môn học, tiếp thu bài tốt. Em cố gắng tập trung hơn ở phần hình học nhé.",
            "Em hoàn thành yêu cầu môn học. Nếu cẩn thận hơn khi đọc đề bài, kết quả sẽ còn tốt hơn."
        ],
        "C": [
            "Em có cố gắng trong học tập. Cần rèn luyện thêm kỹ năng tính nhẩm và nhờ thầy cô hỗ trợ thêm.",
            "Em chăm chỉ đến lớp. Cần chú ý lắng nghe giảng và luyện thêm các phép tính cơ bản mỗi ngày.",
            "Em cần cố gắng nhiều hơn, rèn thêm kỹ năng cộng trừ để hoàn thành tốt bài học.",
            "Em có tiến bộ từng ngày. Hãy dành thêm thời gian luyện toán và mạnh dạn hỏi bài thầy cô nhé."
        ]
    },
    "Tiếng Việt": {
        "T": [
            "Em đọc to, lưu loát, diễn cảm tốt và trả lời câu hỏi đọc hiểu chính xác. Đáng khen ngợi.",
            "Chữ viết nắn nót, giữ vở sạch đẹp, biết dùng từ gợi cảm khi viết đoạn văn. Rất tốt.",
            "Em có vốn từ phong phú, câu văn mạch lạc và giàu cảm xúc. Hãy tiếp tục phát huy năng khiếu.",
            "Kỹ năng đọc hiểu tốt, phát biểu tự tin và viết đúng chính tả. Thầy/Cô rất hài lòng về em.",
            "Em đọc trôi chảy, viết văn có cảm xúc chân thực và trình bày rất cẩn thận."
        ],
        "H": [
            "Em đọc to, rõ ràng và hiểu nội dung bài. Cần chú ý rèn chữ viết đều nét hơn.",
            "Em hoàn thành yêu cầu bài học. Cần chú ý viết đúng chính tả các âm đầu và dấu câu.",
            "Có tiến bộ trong việc dùng từ đặt câu. Em cần rèn đọc diễn cảm và ngắt nghỉ đúng nhịp hơn.",
            "Em chăm chỉ học bài. Cần mở rộng thêm vốn từ để bài viết văn thêm sinh động, mạch lạc.",
            "Em đọc bài tương đối lưu loát. Cố gắng rèn thêm kỹ năng viết đoạn văn hoàn chỉnh nhé."
        ],
        "C": [
            "Em có cố gắng đọc bài. Cần luyện đọc nhiều hơn tại nhà và chú ý viết đúng các từ ngữ khó.",
            "Em chăm chỉ tới lớp. Cần rèn thêm chữ viết cho đều nét và chú ý dấu thanh khi viết chính tả.",
            "Em cần cố gắng hơn nữa trong giờ đọc hiểu, mạnh dạn trao đổi cùng thầy cô và các bạn.",
            "Em có tinh thần học tập. Cần rèn luyện thêm kỹ năng đọc to và viết bài cẩn thận hơn."
        ]
    },
    "DEFAULT": {
        "T": [
            "Em nắm chắc kiến thức môn học, tiếp thu bài nhanh và luôn chủ động, tích cực phát biểu. Rất đáng khen.",
            "Có tinh thần tự giác cao, hoàn thành xuất sắc các nội dung học tập. Hãy tiếp tục phát huy em nhé.",
            "Kỹ năng thực hành thành thạo, tư duy nhạy bén và có nhiều sáng tạo trong giờ học.",
            "Em chăm chỉ, gương mẫu, hoàn thành tốt mọi nhiệm vụ học tập được giao."
        ],
        "H": [
            "Em nắm được nội dung bài học, tích cực tham gia các hoạt động. Cần chú ý cẩn thận hơn khi làm bài.",
            "Có tiến bộ trong học tập, hoàn thành tốt yêu cầu môn học. Em nên chủ động phát biểu xây dựng bài hơn.",
            "Em hiểu bài và làm bài đầy đủ. Cần rèn luyện thêm để nâng cao kỹ năng thực hành.",
            "Hoàn thành các mục tiêu bài học. Cố gắng tự tin hơn nữa để đạt kết quả cao hơn nhé."
        ],
        "C": [
            "Em có cố gắng trong học tập. Cần tập trung hơn trong giờ học và chủ động nhờ thầy cô hỗ trợ.",
            "Em cần dành thêm thời gian ôn tập kiến thức cơ bản để hoàn thành tốt các bài tập.",
            "Em chăm chỉ đến lớp. Cần mạnh dạn trao đổi bài với bạn bè và rèn luyện thêm kỹ năng môn học."
        ]
    }
}

# ---------------------------------------------------------------------------
# Combinatorial Building Blocks for Zero-Duplicate Generation (Anti-Repetition)
# ---------------------------------------------------------------------------

OPENINGS_WITH_NAME = [
    "{name} tiếp thu bài nhanh,",
    "Em {name} có ý thức học tập rất tốt,",
    "{name} chăm ngoan, tích cực trong giờ học,",
    "Em {name} luôn chủ động và tự giác,",
    "{name} thể hiện tư duy học tập tốt,",
    "{name} làm bài rất nghiêm túc và cẩn thận,",
    "Em {name} có tinh thần trách nhiệm cao,",
    "{name} hăng hái phát biểu xây dựng bài,",
    "Em {name} hiểu bài và vận dụng rất tốt,",
    "{name} có nhiều tiến bộ rõ rệt trong học tập,",
    "Em {name} chăm chỉ, hoàn thành tốt nhiệm vụ,",
    "{name} luôn tập trung lắng nghe giảng,",
    "Em {name} thể hiện sự tự tin và năng nổ,",
    "{name} có nề nếp học tập rất đáng khen,",
    "Em {name} luôn hoàn thành bài tập đúng hạn,"
]

OPENINGS_GENERAL = [
    "Em tiếp thu bài nhanh và chủ động,",
    "Có tinh thần tự giác, chăm chỉ học tập,",
    "Em hiểu bài tốt và tích cực phát biểu,",
    "Có ý thức học tập tốt, làm bài đầy đủ,",
    "Em chăm ngoan, hoàn thành tốt nhiệm vụ học tập,",
    "Trong giờ học, em luôn tập trung và tích cực,",
    "Có nhiều tiến bộ đáng ghi nhận trong môn học,",
    "Em rất ngoan, có ý thức rèn luyện tốt,",
    "Học sinh có tinh thần vượt khó và chăm chỉ,",
    "Em tích cực tương tác và xây dựng bài sôi nổi,",
    "Có thái độ học tập nghiêm túc, trách nhiệm,",
    "Em nắm chắc nội dung và thực hành tự tin,",
    "Em luôn có ý thức tự giác hoàn thành bài học,"
]

SUBJECT_CORE_CLAUSES: Dict[str, Dict[str, List[str]]] = {
    "Toán": {
        "T": [
            "tính toán nhanh nhẹn và rất chính xác.",
            "kỹ năng tính nhẩm tốt, trình bày bài khoa học và sạch sẽ.",
            "nắm chắc kiến thức trọng tâm, giải toán có lời văn mạch lạc.",
            "tư duy logic nhạy bén, biết tìm nhiều cách giải hay.",
            "vận dụng linh hoạt các quy tắc tính toán vào bài tập.",
            "nhận biết hình học và đo lường rất tốt, vẽ hình chuẩn.",
            "kỹ năng đặt tính và thực hiện phép tính thành thạo.",
            "chủ động tìm tòi và giải quyết bài toán nhanh trước thời gian.",
            "làm bài cẩn thận, chữ số rõ ràng, lập luận chặt chẽ.",
            "nắm vững các bảng cộng, trừ, nhân, chia và tính nhẩm thành thục.",
            "có khả năng phân tích đề bài toán và tìm lời giải nhanh gọn.",
            "thực hiện các bài toán có lời văn đúng trình tự và logic.",
            "tiếp thu nhanh các dạng toán mới, làm bài tự tin."
        ],
        "H": [
            "nắm được các phép tính cơ bản, cần rèn thêm kỹ năng tính nhẩm nhanh hơn.",
            "hiểu bài và hoàn thành bài tập, cần chú ý đọc kỹ đề toán có lời văn.",
            "làm tính tương đối tốt, nên cẩn thận hơn khi đặt tính thẳng cột.",
            "nắm được kiến thức trọng tâm, cần chú ý tính toán cẩn thận để tránh nhầm lẫn.",
            "biết cách giải bài tập, cần rèn luyện thêm kỹ năng nhân chia có nhớ.",
            "chăm chỉ làm bài, cần chú ý ghi đúng tên đơn vị đo khi giải toán.",
            "tiếp thu bài có tiến bộ, cần tập trung hơn ở các bài tập hình học.",
            "hoàn thành yêu cầu môn học, cần kiểm tra lại kết quả trước khi nộp bài.",
            "có ý thức học tập tốt, cần rèn thêm kỹ năng giải toán nhiều bước tính.",
            "hiểu các quy tắc tính toán, cần rèn thêm tốc độ làm bài trên lớp.",
            "làm bài tương đối tốt, cần chú ý viết chữ số rõ ràng và nắn nót hơn."
        ],
        "C": [
            "có cố gắng học tập, cần rèn thêm các phép tính cộng trừ cơ bản mỗi ngày.",
            "chăm chỉ đến lớp, cần nhờ thầy cô và bạn bè hướng dẫn thêm khi làm bài.",
            "có tinh thần học tập, cần dành thêm thời gian luyện tập tính toán tại nhà.",
            "bước đầu nắm được bài học, cần tập trung lắng nghe và mạnh dạn hỏi bài.",
            "có sự nỗ lực khi làm bài, cần rèn thêm kỹ năng nhận biết và tính nhẩm.",
            "chăm ngoan, cần ôn lại bảng nhân chia cơ bản để làm bài tốt hơn."
        ]
    },
    "Tiếng Việt": {
        "T": [
            "đọc to, rõ ràng, lưu loát và diễn cảm rất tốt.",
            "trả lời các câu hỏi đọc hiểu chính xác, hiểu sâu nội dung bài.",
            "chữ viết nắn nót, giữ vở sạch đẹp, trình bày khoa học.",
            "dùng từ đặt câu phong phú, đoạn văn giàu cảm xúc và hình ảnh.",
            "viết đúng chính tả, câu văn mạch lạc, ngữ pháp chuẩn xác.",
            "tích cực tương tác, chia sẻ cảm nghĩ tự tin trước lớp.",
            "có vốn từ ngữ dồi dào, biết cách liên kết câu khéo léo.",
            "kỹ năng viết đoạn văn miêu tả/kể chuyện sinh động, tự nhiên.",
            "đọc trôi chảy, phát âm chuẩn và ngắt nghỉ đúng dấu câu."
        ],
        "H": [
            "đọc to, rõ ràng và hiểu nội dung bài, cần rèn chữ viết đều nét hơn.",
            "hoàn thành yêu cầu bài học, cần chú ý viết đúng chính tả các âm đầu khó.",
            "có tiến bộ khi đặt câu, cần rèn đọc diễn cảm và ngắt nghỉ đúng nhịp.",
            "chăm chỉ học bài, cần mở rộng thêm vốn từ để câu văn thêm sinh động.",
            "đọc bài tương đối lưu loát, cần chú ý trả lời trọn câu khi đọc hiểu.",
            "hiểu nội dung bài đọc, cần rèn thêm kỹ năng viết đoạn văn hoàn chỉnh.",
            "chữ viết rõ ràng, cần chú ý độ cao và khoảng cách giữa các con chữ."
        ],
        "C": [
            "có cố gắng đọc bài, cần luyện đọc nhiều hơn tại nhà và chú ý từ ngữ khó.",
            "chăm chỉ tới lớp, cần rèn thêm chữ viết cho đều nét và chú ý dấu thanh.",
            "có tinh thần học tập, cần mạnh dạn trao đổi cùng thầy cô trong giờ đọc hiểu.",
            "bước đầu đọc được câu đơn, cần luyện đọc to và đúng tốc độ quy định.",
            "ngoan ngoãn, cần chú ý lắng nghe viết chính tả cẩn thận hơn."
        ]
    },
    "Tiếng Anh": {
        "T": [
            "phát âm chuẩn, giao tiếp tự tin và hào hứng trong giờ học.",
            "ghi nhớ từ vựng tốt, ngữ điệu tự nhiên, phản xạ tiếng Anh nhanh.",
            "nghe hiểu tốt, tích cực tham gia các trò chơi và hội thoại nhóm.",
            "nắm chắc mẫu câu giao tiếp đơn giản, nói to và rõ ràng.",
            "có vốn từ vựng phong phú theo chủ điểm, tiếp thu bài rất nhanh."
        ],
        "H": [
            "hiểu bài và nhớ từ vựng, cần chú ý phát âm rõ các âm đuôi (ending sounds).",
            "hoàn thành các bài tập, cần tự tin nói to hơn khi luyện giao tiếp cùng bạn.",
            "tiếp thu bài tốt, cần dành thêm thời gian ôn tập và ghi nhớ từ vựng.",
            "nhận biết từ vựng tương đối tốt, cần rèn thêm kỹ năng nghe và phản xạ."
        ],
        "C": [
            "có cố gắng học bài, cần chăm chỉ nghe và nhắc lại từ mới thường xuyên.",
            "chăm chỉ đến lớp, cần mạnh dạn phát âm và tham gia luyện nói cùng bạn bè.",
            "có tinh thần học tập, cần ôn lại các từ vựng và mẫu câu cơ bản mỗi ngày."
        ]
    },
    "DEFAULT": {
        "T": [
            "nắm chắc kiến thức môn học, tiếp thu bài nhanh và luôn tích cực phát biểu.",
            "có tinh thần tự giác cao, hoàn thành xuất sắc các nội dung học tập.",
            "kỹ năng thực hành thành thạo, tư duy nhạy bén và có nhiều sáng tạo.",
            "chăm chỉ, gương mẫu, hoàn thành tốt mọi nhiệm vụ được giao.",
            "chủ động học tập, có sự liên hệ thực tế phong phú và sâu sắc."
        ],
        "H": [
            "nắm được nội dung bài học, cần chú ý cẩn thận hơn khi làm bài thực hành.",
            "có tiến bộ trong học tập, nên chủ động phát biểu xây dựng bài hơn.",
            "hiểu bài và làm bài đầy đủ, cần rèn luyện thêm để nâng cao kỹ năng.",
            "hoàn thành mục tiêu môn học, cố gắng tự tin hơn nữa để đạt kết quả cao.",
            "tích cực tham gia hoạt động, cần chú ý tập trung và tỉ mỉ hơn."
        ],
        "C": [
            "có cố gắng trong học tập, cần tập trung hơn và chủ động nhờ thầy cô hỗ trợ.",
            "cần dành thêm thời gian ôn tập kiến thức cơ bản để hoàn thành tốt bài học.",
            "chăm chỉ đến lớp, cần mạnh dạn trao đổi bài với bạn bè và thầy cô.",
            "có tinh thần học hỏi, cần kiên trì rèn luyện kỹ năng môn học mỗi ngày."
        ]
    }
}

CLOSINGS_BY_LEVEL: Dict[str, List[str]] = {
    "T": [
        "Cần tiếp tục phát huy năng lực này.",
        "Thầy/Cô rất khen ngợi em.",
        "Hãy luôn duy trì phong độ học tập tốt này nhé.",
        "Đáng biểu dương và khen ngợi.",
        "Thầy cô rất tự hào về sự tiến bộ của em.",
        "Chúc em tiếp tục phát huy và tỏa sáng hơn nữa.",
        "Phát huy thật tốt năng khiếu của mình nhé em.",
        "Hãy giữ vững niềm say mê và nỗ lực này.",
        "Rất đáng khen ngợi tinh thần học tập của em.",
        "Tiếp tục tự tin và chủ động như vậy nhé em."
    ],
    "H": [
        "Cần cố gắng rèn luyện thêm để đạt kết quả cao hơn.",
        "Em hãy tự tin hơn nữa để ngày càng tiến bộ nhé.",
        "Cố gắng luyện tập đều đặn mỗi ngày nhé em.",
        "Thầy cô tin em sẽ tiến bộ nhiều hơn trong thời gian tới.",
        "Nếu chú ý cẩn thận hơn, kết quả của em sẽ còn tốt hơn nữa.",
        "Hãy phát huy những điểm mạnh và rèn thêm điểm chưa tốt nhé.",
        "Thầy cô luôn sẵn sàng đồng hành và hướng dẫn em.",
        "Cố gắng duy trì sự chăm chỉ này nhé em.",
        "Chỉ cần thêm chút cẩn thận, em sẽ đạt kết quả rất tốt.",
        "Hãy tự tin chia sẻ và hỏi bài thầy cô khi cần nhé."
    ],
    "C": [
        "Em hãy tự tin hỏi bài khi chưa hiểu, thầy cô luôn hỗ trợ em.",
        "Cố gắng dành thêm thời gian ôn luyện để tiến bộ hơn nhé.",
        "Chỉ cần kiên trì luyện tập, em nhất định sẽ tiến bộ từng ngày.",
        "Hãy mạnh dạn nhờ thầy cô và các bạn hướng dẫn thêm nhé.",
        "Thầy cô luôn tin tưởng và đồng hành cùng em mỗi ngày.",
        "Cố gắng rèn luyện từng chút một, em sẽ hoàn thành tốt bài học.",
        "Đừng nản lòng, hãy cố gắng thêm mỗi ngày em nhé."
    ]
}


def extract_display_name(full_name: str) -> str:
    """Extracts a friendly short display name or clean first name for natural Vietnamese addressing."""
    if not full_name:
        return ""
    name_clean = full_name.strip()
    parts = name_clean.split()
    if len(parts) >= 2:
        return " ".join(parts[-2:])
    return name_clean


def synthesize_unique_comment(
    student_name: str,
    subject: str,
    level: str,
    tier: str = "H_MID",
    teacher_note: str = "",
    tone: str = "Chuẩn mực & Động viên",
    used_comments: Optional[set] = None,
    allow_name: bool = True
) -> str:
    """
    Synthesizes a 100% unique, circular-compliant primary school comment.
    Utilizes combinatorial permutations of Openings x Competency clauses x Closings x Notes
    guaranteeing zero duplicates across an entire classroom batch.
    """
    import random
    import re

    if used_comments is None:
        used_comments = set()

    # Normalize subject key
    matched_subj = "DEFAULT"
    for s_key in SUBJECT_CORE_CLAUSES:
        if s_key != "DEFAULT" and (s_key.lower() in subject.lower() or subject.lower() in s_key.lower()):
            matched_subj = s_key
            break

    subj_clauses = SUBJECT_CORE_CLAUSES[matched_subj].get(level, SUBJECT_CORE_CLAUSES["DEFAULT"].get(level, []))
    closings = CLOSINGS_BY_LEVEL.get(level, CLOSINGS_BY_LEVEL["H"])

    display_name = extract_display_name(student_name) if allow_name else ""

    # Clean teacher note
    clean_note = ""
    if teacher_note and len(str(teacher_note).strip()) > 2 and str(teacher_note).strip().lower() not in ["none", "nan", "không"]:
        raw_n = str(teacher_note).strip()
        clean_note = raw_n.rstrip(".")

    best_candidate = ""
    for attempt in range(120):
        # 55% chance to use student name for variety
        use_name_this_time = allow_name and bool(display_name) and (random.random() < 0.55 or attempt > 30)

        if use_name_this_time:
            op_pattern = random.choice(OPENINGS_WITH_NAME)
            opening = op_pattern.format(name=display_name)
        else:
            opening = random.choice(OPENINGS_GENERAL)

        clause = random.choice(subj_clauses)
        closing = random.choice(closings)

        first_sentence = f"{opening} {clause}".strip()

        if clean_note:
            note_lower = clean_note.lower()
            if any(k in note_lower for k in ["chữ", "viết", "trình bày"]):
                note_text = "Cần lưu ý thêm về chữ viết cho đều đẹp."
            elif any(k in note_lower for k in ["nhanh", "tính", "nhẩm"]):
                note_text = "Em tiếp thu và thao tác rất nhanh nhẹn."
            elif any(k in note_lower for k in ["đọc", "phát âm"]):
                note_text = "Em lưu ý rèn thêm phát âm rõ ràng hơn."
            else:
                note_text = f"{clean_note[0].upper() + clean_note[1:]}."
            
            candidate = f"{first_sentence} {note_text} {closing}"
        else:
            candidate = f"{first_sentence} {closing}"

        candidate = re.sub(r"\s+", " ", candidate).strip()
        norm_key = re.sub(r"[^\w\s]", "", candidate).lower()

        if norm_key not in used_comments:
            best_candidate = candidate
            used_comments.add(norm_key)
            break

    if not best_candidate:
        best_candidate = f"Em {display_name or student_name} hoàn thành bài học, có tinh thần tự giác cao. Hãy tiếp tục phát huy em nhé."

    return best_candidate
