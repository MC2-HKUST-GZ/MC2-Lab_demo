import type { Locale } from '../types/common';
import type { HomeAboutCopy, HomeContactLink, HomeResearchFocusItem } from '../types/home';

export interface HomeText {
  newsLabel: string;
  newsAria: string;
}

export const homeText: Record<Locale, HomeText> = {
  en: {
    newsLabel: 'News',
    newsAria: 'Browse latest news'
  },
  'zh-CN': {
    newsLabel: '新闻动态',
    newsAria: '查看最新新闻'
  }
};

export const homeAboutCopy: HomeAboutCopy = {
  intro: 'As the Metaverse continues to evolve, its significance in shaping social interactions, educational experiences, and collaborative opportunities is becoming increasingly essential. This digital frontier holds the potential to create prosocial environments where individuals can connect, learn, and innovate, transcending geographical boundaries and fostering inclusivity. By leveraging immersive technologies, the Metaverse can enhance engagement, promote creativity, and facilitate knowledge sharing, ultimately contributing to a more connected and compassionate society.',
  leaderPrefix: 'Led by ',
  leaderName: 'Professor Pan Hui',
  leaderHref: 'https://panhui.people.ust.hk/',
  leaderSuffix: '—an International Fellow of the Royal Academy of Engineering, Member of Academia Europaea, and IEEE Fellow—the Center for Metaverse and Computational Creativity (MC2) operates at the intersection of immersive technology, Human-Computer Interaction (HCI), social computing, computational social science, and machine learning. Our research is guided by five key foci:',
  closing: 'Our group’s work has been published in top-tier venues across computer science and art, including ACM WWW, ACM SIGCOMM, ACM Mobisys, ACM MobiCom, ACM CoNext, IEEE Infocom, IEEE PerCom, IEEE ICNP, IEEE ICDCS, IJCAI, AAAI, SIGGRAPH, CHI, CSCW, and more. At MC2, we are committed to harnessing these interdisciplinary approaches to create transformative experiences that enrich lives and foster a more connected world. Join us as we navigate the exciting possibilities of the Metaverse and its applications in society.'
};

export const homeResearchFocus: HomeResearchFocusItem[] = [
  {
    title: 'Pushing the Boundaries of Immersive Technologies',
    body: 'We explore the convergence of virtual reality (VR), augmented reality (AR), extended reality (XR), and mixed reality (MR), alongside other transformative technologies that shape the Metaverse.'
  },
  {
    title: 'Social Computing and Prosocial Interactions ',
    body: 'We investigate how social computing can enhance prosocial behaviors, fostering trust, engagement, and collaboration within virtual communities.'
  },
  {
    title: 'Generative AI and Large Language Models Techniques',
    body: 'As core technologies driving our research, we focus on developing applicable generative AI and LLMs-related techniques for enhancing the interactivity in metaverse and broadening the boundaries of higher education.'
  },
  {
    title: 'VR/AR/XR Computer Network ',
    body: 'We examine innovative approaches on VR-related network usage that adapt to user needs and environmental conditions, ensuring seamless experiences in the Metaverse.'
  },
  {
    title: 'AI+ Projects ',
    body: 'Our AI+ initiative, particularly in education, encompasses three main components: Metaverse classrooms, AI lecturers, and the development of interactive AI lecturers. These projects aim to create immersive, cross-campus learning environments, introduce AI-driven teaching methods, and develop real-time interactive AI lecturers, thereby revolutionizing the educational landscape.'
  }
];

export const homeContactLinks: HomeContactLink[] = [
  {
    label: 'WeChat',
    href: '#'
  },
  {
    label: 'YouTube',
    href: 'https://www.youtube.com/@MC2HKUSTGZCWB'
  },
  {
    label: 'Medium',
    href: 'https://medium.com/@mc2.hkust.hkustgz'
  }
];
