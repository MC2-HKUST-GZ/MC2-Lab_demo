(function () {
  const recentPublicationData = {
    2026: [
      {
        title: `IMGNN: An Efficient, Effective and Generalizable Algorithm for Influence Maximization in Social Networks.`,
        authors: `Haotian Zhang, Kai Han, Zhizhuo Yin, Shuang Cui, Jing Tang, and Pan Hui.`,
        venue: `In Proceedings of the 32nd ACM SIGKDD Conference on Knowledge Discovery and Data Mining (KDD), Toronto, Canada, August 2026.`
      },
      {
        title: `Co-Designing Digital Humans for Online Learning: A Framework for Human-AI Pedagogical Integration.`,
        authors: `Xiaokang Lei, Christie Pang, Yuyang Jiang, Tong Xin, and Pan Hui.`,
        venue: `In Proceedings of the 28th Human-Computer Interaction International Conference (HCII 2026), Montreal, Canada, July 2026.`
      },
      {
        title: `Cross-Platform Short-Video Diplomacy: Topic and Sentiment Analysis of China-US Relations on Douyin and TikTok.`,
        authors: `Zheng Wei, Mingchen Li, Junxiang Liao, Zeyu Yang, Xiaoyu Yang, Yixuan Xie, Pan Hui, and Huamin Qu.`,
        venue: `In Proceeding of the International AAAI Conference on Web and Social Media (ICWSM 2026), Los Angeles, USA, June 2026.`
      },
      {
        title: `Unpacking Discourses on Childbirth and Parenthood in Popular Social Media Platforms Across China, Japan, and South Korea.`,
        authors: `Zheng Wei, Yunqi Li, Yucheng He, Yuelu Li, Xian Xu, Huamin Qu, Pan Hui, and Muzhi Zhou.`,
        venue: `In Proceeding of the International AAAI Conference on Web and Social Media (ICWSM 2026), Los Angeles, USA, June 2026.`
      },
      {
        title: `FovRL: Joint Foveation and Quality Control for Immersive VR Streaming Using Reinforcement Learning.`,
        authors: `Yuk Hang Tsui, Ze Wu, Ahmad Alhilal, Matti Siekkinen, and Pan Hui.`,
        venue: `In Proceedings of the ACM Web Conference 2026 (WWW '26), Dubai, United Arab Emirates, April 13-17, 2026.`
      },
      {
        title: `Collaborative Social Bot Detection in the Fediverse.`,
        authors: `Min Gao, Wen Wen, Haoran Du, Qiang Duan, Yu Xiao, Yupeng Li, Xin Wang, Pan Hui and Yang Chen.`,
        venue: `In Proceedings of the ACM Web Conference 2026 (WWW '26), Dubai, United Arab Emirates, April 13-17, 2026.`
      },
      {
        title: `CityLens: Evaluating Large Vision-Language Models for Urban Socioeconomic Sensing.`,
        authors: `Tianhui Liu, Hetian Pang, Xin Zhang, Tianjian Ouyang, Zhiyuan Zhang, Jie Feng, Yong Li, and Pan Hui.`,
        venue: `In The Fourteenth International Conference on Learning Representations (ICLR '26), Rio de Janeiro, Brazil, April, 2026.`
      },
      {
        title: `SesaHand: Enhancing 3D Hand Reconstruction via Controllable Generation with Semantic and Structural Alignment.`,
        authors: `Zhuoran Zhao, Xianghao Kong, Linlin Yang, Wei Zheng, Pan Hui, Anyi Rao.`,
        venue: `In The Fourteenth International Conference on Learning Representations (ICLR '26), Rio de Janeiro, Brazil, April, 2026.`
      },
      {
        title: `Where Digital Meets Place: Deriving Strategies for Curating Mixed Reality Exhibitions in Public Spaces.`,
        authors: `Yawei Zhao, Jiaxin Liang, Hao Li, and Pan Hui.`,
        venue: `In Proceedings of the 2026 CHI Conference on Human Factors in Computing Systems (CHI '26), Barcelona, Spain, April, 2026.`
      },
      {
        title: `FAIR: Framing AI's Role in Programming Competitions - Understanding How LLMs Are Changing the Game in Competitive Programming.`,
        authors: `Dongyijie Primo Pan, Lan Luo, Ji Zhu, Zhiqi Gao, Xin Tong, and Pan Hui.`,
        venue: `In Proceedings of the 44th ACM Conference on Human Factors in Computing Systems (CHI '26), Barcelona, Spain, April 2026.`
      },
      {
        title: `LingoLift: Supporting Educators in Personalized Oral Language Teaching for Autistic Children through Content Generation.`,
        authors: `Jiawen Zhang, Dongyijie Primo Pan, Li Wang, Pan Hui, and Xin Tong.`,
        venue: `In Proceedings of the 44th ACM Conference on Human Factors in Computing Systems (CHI '26), Barcelona, Spain, April 2026.`
      },
      {
        title: `Overcoming Translation Delays: Towards Better Subtitle Design for Foreign Language Conversations in Extended Reality.`,
        authors: `Ziming Li, Rongkai Shi, Hongji Li, Jialin Wang, Pan Hui, and Hai-Ning Liang.`,
        venue: `In Proceedings of the 2026 CHI Conference on Human Factors in Computing Systems (CHI '26), Barcelona, Spain, April 2026.`
      },
      {
        title: `GuideMe: A VLM-Based System Assisting Independent Smartphone Learning for Older Adults.`,
        authors: `Kairong Fang, Jiesi Zhang, Shi-Ting Ni, Pan Hui, and Yuyang Wang.`,
        venue: `In Proceedings of the 2026 CHI Conference on Human Factors in Computing Systems (CHI '26), Barcelona, Spain, April 2026.`
      },
      {
        title: `Hidden Labor behind the Hype: Understanding AI Side Hustles through Platform Narratives and Worker Practices.`,
        authors: `Xiaoyu Yang, Zelin Zhao, Weipeng Chen, Corey Kewei Xu, and Pan Hui.`,
        venue: `In Proceedings of the 2026 CHI Conference on Human Factors in Computing Systems (CHI '26), Barcelona, Spain, April 2026.`
      },
      {
        title: `When Effort Becomes Visible: Facet-Level Shifts in Evaluation and Workload during VR Teamwork.`,
        authors: `Zheng Wei, Hyeonmin Lee, Yawei Zhao, Junxiang Liao, Hao Li, Xian Xu, Hayoung Oh, Lik-Hang Lee, Wai Tong, and Pan Hui.`,
        venue: `In Proceedings of the IEEE Conference on Virtual Reality and 3D User Interfaces (IEEE VR '26), Daegu, Korea, March, 2026 (Selected for publication at IEEE Transactions on Visualization and Computer Graphics (TVCG)).`
      },
      {
        title: `Saliency-Guided Foveated Video Encoding for Low-Latency and Immersive Cloud VR.`,
        authors: `Ze Wu, Ahmad Alhilal, Yuk Hang Tsui, Wenjye Chai, Matti Siekkinen, and Pan Hui.`,
        venue: `In Proceedings of the IEEE Conference on Virtual Reality and 3D User Interfaces (IEEE VR '26), Daegu, Korea, March, 2026 (Selected for publication at IEEE Transactions on Visualization and Computer Graphics (TVCG)).`
      },
      {
        title: `RayFlex: Inducing Weight Perception through Raycast Pseudo-Haptics in Virtual Reality.`,
        authors: `Yushi Wei, Sen Zhang, Rongkai Shi, Simon Fong, Pan Hui, and Hai-Ning Liang.`,
        venue: `In Proceedings of the IEEE Conference on Virtual Reality and 3D User Interfaces (IEEE VR '26), Daegu, Korea, March, 2026 (Selected for publication at IEEE Transactions on Visualization and Computer Graphics (TVCG)).`
      },
      {
        title: `Aligning Gamification with Learner Motivation: Insights from VR-Based Learning Tasks.`,
        authors: `Yuxuan Bai, Yuzhu Wang, Pan Hui, and Yuyang Wang.`,
        venue: `In Proceedings of the IEEE Conference on Virtual Reality and 3D User Interfaces (IEEE VR '26), Daegu, Korea, March, 2026 (Selected for publication at IEEE Transactions on Visualization and Computer Graphics (TVCG)).`
      },
      {
        title: `Enhancing Creativity in Virtuality: How Annotations in Creative Support Tool Innovate Design Ideation in Virtual Reality.`,
        authors: `Xuetong Wang, Ching Christie Pang, Zhen Wu, Ahmad Alhilal, Simin Yang, Tristan Braud, and Pan Hui.`,
        venue: `In Proceedings of the IEEE Conference on Virtual Reality and 3D User Interfaces (IEEE VR '26), Daegu, Korea, March, 2026.`
      },
      {
        title: `Frame Complexity-Aware Foveated Video Encoding for Real-time High-Quality Streaming.`,
        authors: `Ze Wu, Ahmad Alhilal, Yuk Hang Tsui, Matti Siekkinen, and Pan Hui.`,
        venue: `In Proceedings of the IEEE Conference on Virtual Reality and 3D User Interfaces (IEEE VR '26), Daegu, Korea, March, 2026.`
      },
      {
        title: `FlexiCamAR: Enhancing Everyday Camera Interactions on AR Glasses with a Flexible Additional Viewpoint.`,
        authors: `Ziming Li, Hongji Li, Jialin Wang, Pan Hui, and Hai-Ning Liang.`,
        venue: `In Proceedings of the IEEE Conference on Virtual Reality and 3D User Interfaces (IEEE VR '26), Daegu, Korea, March, 2026.`
      },
      {
        title: `FSWAN: Spectrometry-based Wearable Biosensing for Monitoring Plastic Particles in the Human Body.`,
        authors: `Kevin Post, Peo Kuchida, Zhigang Yin, Pan Hui, Petteri Nurmi, and Huber Flores.`,
        venue: `In Proceedings of the ACM International Workshop on Mobile Computing Systems and Applications (HotMobile 2026), Atlanta, Georgia, USA, February, 2026.`
      },
      {
        title: `LLM Safety in Judicial AI: A Stress Test of Social Media Influence on Real-World Judgments.`,
        authors: `Yixuan Xie, Yang He, Xiaoyu Yang, Xu Gai, and Pan Hui.`,
        venue: `In Proceedings of the 40th AAAI Conference on Artificial Intelligence (AAAI '26), Singapore, Jan, 2026.`
      },
      {
        title: `When Trust Collides: Exploring Human-LLM Cooperation Intention through the Prisoner's Dilemma.`,
        authors: `Guanxuan Jiang, Shirao Yang, Yuyang Wang and Pan Hui.`,
        venue: `In International Journal of Human-Computer Studies (IJHCS), Jan., 2026.`
      },
      {
        title: `The Geography of Climate Concern: A Large-scale Analysis of Public Discourse on Extreme Heat in China using Social Media and Explainable AI.`,
        authors: `Pu Zhang, Yiliang Li, Zheng Wei, and Pan Hui.`,
        venue: `In Environmental Impact Assessment Review, 117, 108227, 2026.`
      },
      {
        title: `Leveraging large language models to explain provincial heterogeneity in online discussion of water pollution in China: A socioeconomic and environmental perspective.`,
        authors: `Pu Zhang, Na Jiang, Yiliang Li, Zheng Wei, and Pan Hui.`,
        venue: `In Journal of Cleaner Production, 543, 147517, 2026.`
      },
      {
        title: `Beyond Aggregate Feedback: A Spatially-Aware Computational Framework for Understanding Citizen-State Interaction on Social Media.`,
        authors: `Pu Zhang, Zheng Wei, Muzhi Zhou, James Evans, and Pan Hui.`,
        venue: `In ACM Transactions on Social Computing, 2026.`
      }
    ],
    2025: [
      {
        title: `EyeNexus: Adaptive Gaze-Driven Quality and Bitrate Streaming for Seamless VR Cloud Gaming Experiences.`,
        authors: `Ze Wu, Ahmad Alhilal, Yuk Hang Tsui, Matti Siekkinen, and Pan Hui.`,
        venue: `In Proceedings of the 20th International Conference on emerging Networking EXperiments and Technologies (CoNEXT), Hong Kong, December, 2025.`
      },
      {
        title: `PyraMotion: Attentional Pyramid-Structured Motion Integration for Co-Speech 3D Gesture Synthesis.`,
        authors: `Zhizhuo Yin, Yuk Hang Tsui, and Pan Hui.`,
        venue: `In Proceedings of the Thirty-Ninth Annual Conference on Neural Information Processing Systems (NeurIPS), San Diego, CA, USA, December 2025.`
      },
      {
        title: `" Is This Your True Intention to Participate?"?Simulated Interrogation and Hidden Control in an Agent-Driven Installation.`,
        authors: `Yuyang Jiang, Luwen Yu, Jia Sun, Jiashun Chen, Pui Wa Chau, and Pan Hui.`,
        venue: `In Proceedings of the SIGGRAPH Asia 2025 Art Papers, Hong Kong, December 2025.`
      },
      {
        title: `Revolutionizing Education: Integrating Metaverse and the First AI Lecturers in Cross-Campus Higher Education.`,
        authors: `Ching Christie Pang, Yawei Zhao, Zhizhuo Yin, Jia Sun, Yuk Hang Tsui, Shan Jin, and Pan Hui.`,
        venue: `In SIGGRAPH Asia 2025 Educator's Forum (SA '25), Article 24, 1-9, 2025 (Best Education Paper at SIGGRAPH Asia 2025, the work was also covered by Nature).`
      },
      {
        title: `PreciseTouch: A Multi-Finger Technique for Targeting Small Items on Touchscreens.`,
        authors: `Zhonghua Sheng, Simin Yang, Xiaofu Jin, Lik-Hang Lee, Pan Hui, and Huamin Qu.`,
        venue: `In International Journal of Human-Computer Interaction (IJHCI) (Publication date 2025/12/13).`
      },
      {
        title: `SARF: Sparsity-Aware Reconstruction Framework for Large-Scale Datasets.`,
        authors: `ZAgustin Zuniga, Huber Flores, Ngoc Thi Nguyen, Pan Hui, Sasu Tarkoma, and Petteri Nurmi.`,
        venue: `In IEEE Transactions on Big Data (Publication date 2025/12/4).`
      },
      {
        title: `SemiPoint: Generalizing Cross-Scene Point Cloud Video Streaming with Semi-Supervised Learning and Residual-Aware Adaptation.`,
        authors: `Ao Yu, Yang Hui, Cuiyang Feng, Petteri Nurmi, Mohamed Cheriet, and Pan Hui.`,
        venue: `In IEEE Transactions on Multimedia (Publication date 2025/12).`
      },
      {
        title: `Vetaverse: Technologies, Applications, and Visions toward the Intersection of Metaverse, AI, Vehicles, and Transportation Systems.`,
        authors: `Peng Yuan Zhou, Jinjing Zhu, Yiting Wang, Yunfan Lu, Zixiang Wei, Haolin Shi, Yuchen Ding, Yu Gao, Qinglong Huang, Yan Shi, Ahmad Alhilal, Lik-Hang Lee, Tristan Braud, Pan Hui, and Lin Wang.`,
        venue: `In IEEE Transactions on Artificial Intelligence (Publication date 2025/11/25).`
      },
      {
        title: `CGM-Led Multimodal Tracking with Chatbot Support: An Autoethnography in Sub-Health.`,
        authors: `Dongyijie Primo Pan, Lan Luo, Yike Wang, and Pan Hui.`,
        venue: `In Proceedings of the International Conference on Human-Engaged Computing (ICHEC 2025), Singapore, November 2025 (Best Short Paper Award).`
      },
      {
        title: `Generative Multi-Sensory Meditation: Exploring Immersive Depth and Activation in Virtual Reality.`,
        authors: `Yuyang Jiang, Binzhu Xie, Lina Xu, Xiaokang Lei, Shi Qiu, Luwen Yu, and Pan Hui.`,
        venue: `In Proceedings of the 33rd ACM International Conference on Multimedia (MM '25), Dublin, Ireland, October, 2025.`
      },
      {
        title: `My Dataset of Love: A Preliminary Mixed-Method Exploration of Human-AI Romantic Relationships.`,
        authors: `Xuetong Wang, Ching Christie Pang, and Pan Hui.`,
        venue: `In Proceedings of the ACM on Human-Computer Interaction (CSCW), Bergen, Norway, October 2025 (Honourable Mention, and referred by Japan Times).`
      },
      {
        title: `DreamScene: 3D Gaussian-based End-to-end Text-to-3D Scene Generation.`,
        authors: `Haoran Li, Yuli Tian, Kun Lan, Yong Liao, Lin Wang, Pan Hui, and Peng Yuan Zhou.`,
        venue: `In IEEE Transactions on Pattern Analysis and Machine Intelligence (Publication date 2025/10/15).`
      },
      {
        title: `Exploring Gaze Dynamics in VR Film Education: Gender, Avatar, and the Shift Between Male and Female Perspectives.`,
        authors: `Zheng Wei, Jia Sun, Junxiang Liao, Lik-Hang Lee, Pan Hui, Huamin Qu, Wai Tong, Xian Xu.`,
        venue: `In IEEE International Symposium on Mixed and Augmented Reality (ISMAR 2025), Daejeon, Korea, October 2025.`
      },
      {
        title: `Talking Spell: A Wearable System Enabling Real-Time Anthropomorphic Voice Interaction with Everyday Objects.`,
        authors: `Xuetong Wang, Ching Christie Pang, and Pan Hui.`,
        venue: `In Proceedings of the 38th Annual ACM Symposium on User Interface Software and Technology (UIST '25), Seoul, Republic of Korea, September 2025.`
      },
      {
        title: `CineVision: An Interactive Pre-visualization Storyboard System for Director-Cinematographer Collaboration.`,
        authors: `Zheng Wei, Hongtao Wu, lvmin Zhang, Xian Xu, Yefeng Zheng, Pan Hui, Maneesh Agrawala, Huamin Qu, Anyi Rao.`,
        venue: `In Proceeding of the ACM Symposium on User Interface Software and Technology (UIST 2025), Busan Korea, September 2025.`
      },
      {
        title: `MetaCineMoji: Visualizing film set communication in an interactive interface for collaboration in virtual LED production.`,
        authors: `Zheng Wei, Shan Jin, Wai Tong, Pan Hui, Lik-Hang Lee, and Xian Xu.`,
        venue: `In Visual Informatics (Publication date 2025/9/25).`
      },
      {
        title: `Little Hands, Big Exhibitions ? Children's Virtual Curation with Generative MR Toolbox.`,
        authors: `Yuyang Jiang, Cekai Weng, Ching Christie Pang, Pengfei Yu, Yuqi Liang, and Pan Hui.`,
        venue: `In Proceedings of the Conference on Animation and Interactive Art (Expanded 2025), Linz, Austria, September, 2025.`
      },
      {
        title: `Algorithmic Miner: Humanity in Service-An AI-Driven VR Journey into Machine Logic.`,
        authors: `Jia Sun, Zheng Wei, and Pan Hui.`,
        venue: `In Proceedings of the ACM on Computer Graphics and Interactive Techniques (SIGGRAPH), Vancouver Canada, August 2025 (Art Paper, Highlighted by SIGGRAPH 2025 Press Release).`
      },
      {
        title: `Sway: Too Numerous to Record.`,
        authors: `Yuyang Jiang, Luwen Yu, Xiaokang Lei, Lei Huang, and Pan Hui.`,
        venue: `In Proceedings of the Special Interest Group on Computer Graphics and Interactive Techniques (SIGGRAPH) Conference Art Gallery, Vancouver, Canada, August 2025.`
      },
      {
        title: `Progressive Speech-Driven VR: Merging Technical Innovation with Artistic Explorations of Identity.`,
        authors: `Yishuai Zhang, and Pan Hui.`,
        venue: `In Proceedings of EVA London 2025, London UK, July 2025.`
      },
      {
        title: `Analyzing the Synthetic-to-Real Domain Gap in 3D Hand Pose Estimation.`,
        authors: `Zhuoran Zhao, Linlin Yang, Pengzhan Sun, Pan Hui, Angela Yao.`,
        venue: `In Proceedings of the Computer Vision and Pattern Recognition (CVPR) Conference, Nashville USA, June 2025.`
      },
      {
        title: `Artificial Human Lecturers: Initial Findings from Asia's First AI Lecturers in Class to Promote Innovation in Education.`,
        authors: `Ching Christie Pang, Yawei Zhao, Zhizhuo Yin, Jia Sun, Reza Hadi Mogavi, and Pan Hui.`,
        venue: `In Proceedings of the 27th HCI International Conference (HCII 2025), Gothenburg Sweden, June 2025.`
      },
      {
        title: `User-Centered Trustworthy AI Metrics in the Metaverse.`,
        authors: `Peng Yuan Zhou, Benjamin Finley, Lik-Hang Lee, Reza Hadi Mogavi, and Pan Hui.`,
        venue: `In Proceedings of the 27th HCI International Conference (HCII 2025), Gothenburg Sweden, June 2025.`
      },
      {
        title: `Deepfake in the Metaverse: An Outlook Survey.`,
        authors: `Haojie Wu, Yong Liao, Reza Hadi Mogavi, Pan Hui, and Peng Yuan Zhou.`,
        venue: `In Proceedings of the 27th HCI International Conference (HCII 2025), Gothenburg Sweden, June 2025.`
      },
      {
        title: `User-Centered Trustworthy AI Metrics in the Metaverse.`,
        authors: `Reza Hadi Mogavi, and Pan Hui.`,
        venue: `In HCI in Games: 7th International Conference, HCI-Games 2025, Held as Part of the 27th HCI International Conference (HCII 2025), Gothenburg Sweden, June 2025.`
      },
      {
        title: `Comments Pusher.`,
        authors: `Yuyang Jiang, Xiaofan Ruan, Wenqiu Xu, Yi Zhang, Xiong Gao, and Pan Hui.`,
        venue: `In ISEA2025 Seoul Creators' Universe (pp. 18-19). 30th International Symposium on Electronic/Emerging Art, Seoul, Republic of Korea, May 2025 (Art Gallery, Acceptance rate 11%).`
      },
      {
        title: `EARVP: Efficient Aggregation for Federated Learning with Robustness, Verifiability, and Privacy.`,
        authors: `Caihong Yan, Xiaofeng Lu, Pietro Lio, and Pan Hui.`,
        venue: `In IEEE Transactions on Information Forensics and Security (Date of Acceptance: 22 May 2025).`
      },
      {
        title: `Illuminating the scene: How virtual environments and learning modes shape film lighting mastery in virtual reality.`,
        authors: `Zheng Wei, Jia Sun, Junxiang Liao, Lik-Hang Lee, Chan In Sio, Pan Hui, Huamin Qu, Wai Tong, and Xian Xu.`,
        venue: `In IEEE Transactions on Visualization and Computer Graphics (Volume: 31, Issue: 5, May 2025) (Best Paper Honorable Mention at IEEE VR 2025 Conference).`
      },
      {
        title: `Toward Personalizable AI Node Graph Creative Writing Support: Insights on Preferences for Generative AI Features and Information Presentation Across Story Writing Processes.`,
        authors: `Hua Xuan Qin, Guangzhi Zhu, Mingming Fan, and Pan Hui.`,
        venue: `In Proceedings of the 2025 CHI Conference on Human Factors in Computing Systems (CHI), Yokohama Japan, April 2025.`
      },
      {
        title: `Exploring LLM-Powered Role and Action-Switching Pedagogical Agents for History Education in Virtual Reality.`,
        authors: `Zihao Zhu, Ao Yu, Xin Tong, and Pan Hui.`,
        venue: `In Proceedings of the 2025 CHI Conference on Human Factors in Computing Systems (CHI), Yokohama Japan, April 2025.`
      },
      {
        title: `“Even When Success Seems Impossible, I Keep Streaming”: How Do Chinese Elderly Streamers Interact with Platform Algorithmic (In)visibility.`,
        authors: `Junxiang Liao, Zheng Wei, Zeyu Yang, Xian Xu, Pan Hui, Changyang He, and Muzhi Zhou.`,
        venue: `In Proceedings of the 2025 CHI Conference on Human Factors in Computing Systems (CHI), Yokohama Japan, April 2025.`
      },
      {
        title: `Parental Perceptions of Children's Deaf Identity Shaping Technology Use: A Qualitative Study on Communication Technologies in Mixed Hearing Families.`,
        authors: `Keyi Zeng, Jingyang Lin, Ruiqi Chen, RAY LC, Pan Hui, and Xin Tong.`,
        venue: `In Proceedings of the Extended Abstracts of the CHI Conference on Human Factors in Computing Systems, Yokohama Japan, April 2025.`
      },
      {
        title: `Interoperability of the metaverse: A digital ecosystem perspective review.`,
        authors: `Liang Yang, Shi-Ting Ni, Yuyang Wang, Ao Yu, Jyh-An Lee, and Pan Hui.`,
        venue: `In IEEE Engineering Management Review (Date of Acceptance: 28 April 2025).`
      },
      {
        title: `Helora: Lora-heterogeneous federated fine-tuning for foundation models.`,
        authors: `Boyu Fan^, Xiang Su, Sasu Tarkoma, and Pan Hui.`,
        venue: `In ACM Transactions on Internet Technology (Date of Acceptance: 25 April 2025).`
      },
      {
        title: `The art of dish: What makes cooked food visually appealing?`,
        authors: `Jiayu Shu, Lik-Hang Lee, Yusen Sun, Pearl Pu, and Pan Hui.`,
        venue: `In Displays (Date of Acceptance: 15 April 2025).`
      },
      {
        title: `EmojiChat: Toward designing emoji-driven social interaction in VR museums.`,
        authors: `Luyao Shen, Xian Wang, Sijia Li, Lik-Hang Lee, Mingming Fan, and Pan Hui.`,
        venue: `In International Journal of Human-Computer Interaction (IJHCI) (Date of Acceptance: 12 April 2025).`
      },
      {
        title: `I Can't Even Recall What I Bought: How Design Influences Impulsive Buying in Douyin Live Sales.`,
        authors: `Zheng Wei, Lik-Hang Lee, Wai Tong, Xian Xu, Chaozhe Zhang, Huamin Qu, and Pan Hui.`,
        venue: `In International Journal of Human-Computer Interaction (IJHCI) (Date of Acceptance: 3 April 2025).`
      },
      {
        title: `Dynamic Hierarchical Reinforcement Learning Framework for Energy-Efficient 5G Base Stations in Urban Environments.`,
        authors: `Dianlei Xu^, Xiang Su, Gopika Premsankar, Huandong Wang, Sasu Tarkoma, and Pan Hui.`,
        venue: `In IEEE Transactions on Mobile Computing (Date of Acceptance: 2 April 2025).`
      },
      {
        title: `Framing metaverse identity: A multidimensional framework for governing digital selves.`,
        authors: `Liang Yang, Yan Xu, and Pan Hui.`,
        venue: `In Telecommunications Policy (Date of Acceptance: 1 April 2025).`
      },
      {
        title: `SNAKE: Harnessing Human Touch for Produce Quality Estimation to Foster Sustainable Retail Practices.`,
        authors: `Zhigang Yin, Marko Radeta, Mohan Liyanage, Mayowa Olapade, Abdul-Rasheed Ottun, Agustin Zuniga, Pan Hui, Petteri Nurmi, and Huber Flores.`,
        venue: `In ACM Transactions on Sensor Networks (Date of Acceptance: 1 April 2025).`
      },
      {
        title: `Exploring user engagement by diagnosing visual guides in onboarding screens with linear regression and XGBoost.`,
        authors: `Lik-Hang Lee, Kit-Yung Lam, and Pan Hui.`,
        venue: `In Displays (Date of Acceptance: 1 April 2025).`
      },
      {
        title: `Toward Sustainable 6G leveraging Digital Twin and Artificial Intelligence: Framework and Case Study.`,
        authors: `Dianlei Xu, Xiang Su, Sasu Tarkoma, and Pan Hui.`,
        venue: `In IEEE Communications Magazine (Date of Acceptance: 12 March 2025).`
      },
      {
        title: `Perceived User Reachability in Mobile UIs Using Data Analytics and Machine Learning.`,
        authors: `Lik-Hang Lee, Yui-Pan Yau, and Pan Hui.`,
        venue: `In International Journal of Human-Computer Interaction (IJHCI) (Date of Acceptance: 16 February 2025).`
      },
      {
        title: `Identity, Crimes, and Law Enforcement in the Metaverse.`,
        authors: `Hua Xuan Qin, Yuyang Wang, and Pan Hui.`,
        venue: `In Nature Humanities and Social Sciences Communications (Date of Acceptance: 12 February 2025).`
      }
    ]
  };

  function escapeHtml(value) {
    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function decoratePanHuiNames(root) {
    root.querySelectorAll('.author').forEach((element) => {
      if (element.innerHTML.includes('pan-hui-name')) {
        return;
      }

      const text = element.textContent || '';
      if (!text.includes('Pan Hui')) {
        return;
      }

      element.innerHTML = escapeHtml(text).replace(
        /Pan Hui/g,
        '<span class="pan-hui-name">Pan Hui</span>'
      );
    });
  }

  function buildPublicationItem(entry) {
    const item = document.createElement('article');
    item.className = 'recent-publication-item';

    const title = document.createElement('h4');
    title.textContent = entry.title;

    const conference = document.createElement('p');
    conference.className = 'conference';
    conference.textContent = entry.venue;

    const author = document.createElement('p');
    author.className = 'author';
    author.textContent = entry.authors;

    item.appendChild(title);
    item.appendChild(conference);
    item.appendChild(author);

    return item;
  }

  function buildYearSection(year, entries) {
    const section = document.createElement('section');
    section.id = `yr${year}`;
    section.className = 'recent-year-section';

    const spacer = document.createElement('br');
    const spacerClone = spacer.cloneNode();
    const title = document.createElement('h3');
    title.textContent = `YEAR ${year}`;

    const list = document.createElement('div');
    list.className = 'recent-publication-list';

    entries.forEach((entry) => {
      list.appendChild(buildPublicationItem(entry));
    });

    section.appendChild(spacer);
    section.appendChild(spacerClone);
    section.appendChild(title);
    section.appendChild(list);

    return section;
  }

  function insertRecentYearNav(navList, years) {
    const firstExistingNav = navList.querySelector('.yr__nav');
    const navFragment = document.createDocumentFragment();

    years.forEach((year) => {
      const navItem = document.createElement('li');
      navItem.className = `yr__nav yr${year}`;

      const anchor = document.createElement('a');
      anchor.href = `#yr${year}`;
      anchor.textContent = year;

      navItem.appendChild(anchor);
      navFragment.appendChild(navItem);
    });

    navList.insertBefore(navFragment, firstExistingNav);

    navList.querySelectorAll('.yr__nav').forEach((navItem, index) => {
      navItem.classList.toggle('active', index === 0);
    });
  }

  function renderRecentPublications() {
    const navList = document.querySelector('.yr__navs');
    const anchorSection = document.getElementById('yr2024');
    const sectionContainer = anchorSection?.parentElement;

    if (!navList || !anchorSection || !sectionContainer) {
      return;
    }

    const years = Object.keys(recentPublicationData)
      .map((year) => Number(year))
      .sort((a, b) => b - a)
      .map(String);

    if (!document.querySelector('.yr__nav.yr2026')) {
      insertRecentYearNav(navList, years);
    }

    years.forEach((year) => {
      if (document.getElementById(`yr${year}`)) {
        return;
      }

      sectionContainer.insertBefore(buildYearSection(year, recentPublicationData[year]), anchorSection);
    });

    decoratePanHuiNames(document);
  }

  renderRecentPublications();
})();
