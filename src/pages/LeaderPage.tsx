import React, { useEffect, useMemo } from 'react';
import SitePageShell from '../components/SitePageShell';
import type { Locale } from '../types/common';
import '../styles/leader-page.css';

interface LeaderPageProps {
  locale?: Locale;
}

const leaderContent = {
  en: {
    pageLabel: 'Director',
    documentTitle: 'Professor Pan Hui | MC2',
    profileAlt: 'Professor Pan Hui',
    title: 'Professor Pan Hui',
    subtitle: 'Director of MC2; Chair Professor and Acting Head of Computational Media and Arts at HKUST(GZ)',
    meta: 'International Fellow of the Royal Academy of Engineering, Member of Academia Europaea, IEEE Fellow, and ACM Distinguished Scientist.',
    homepageLabel: 'Official Homepage',
    sections: [
      {
        title: 'Biography',
        paragraphs: [
          'Professor Pan Hui is Chair Professor and Acting Head of Computational Media and Arts at the Hong Kong University of Science and Technology (Guangzhou), Director of the Center for Metaverse and Computational Creativity (MC2), Chair Professor of Emerging Interdisciplinary Areas at HKUST, Director of the HKUST-DT Systems and Media Laboratory (SyMLab), and Professor of Computer Science at the University of Helsinki. He previously served as Associate Dean of the Information Hub at HKUST(GZ), and was a tenured faculty member in Computer Science and Engineering at HKUST before joining CMA as a founding member.',
          'He studied at The University of Hong Kong and the University of Cambridge, and has built an interdisciplinary research profile spanning mobile computing, computer networking, augmented reality, data science, and human-centered computing. For his sustained contributions to mobile computing and networking, he was elected International Fellow of the Royal Academy of Engineering, Member of Academia Europaea, IEEE Fellow, and ACM Distinguished Scientist. As of 2026, his Google Scholar citations are approaching 40,000.'
        ]
      },
      {
        title: 'Research Agenda',
        paragraphs: [
          "Professor Hui's work can be understood through two connected threads: data-driven systems design and immersive human-data interaction. On one side, he studies human behavior and sociotechnical systems through large-scale empirical data, and translates those insights into software and network systems for mobile, wearable, cloud, and edge environments. On the other side, he develops AR/VR/XR, spatial computing, and metaverse infrastructures that bring digital content, social interaction, and intelligent systems into physical space in more seamless ways."
        ]
      }
    ],
    projectsTitle: 'Selected Directions and Projects',
    projects: [
      {
        title: 'AI-enabled education',
        body: "In 2024, he led the launch of one of Asia's first AI lecturer initiatives, using AI-generated digital humans in teaching and exploring how realism, trust, and pedagogy interact in immersive classrooms. The work later drew broader attention from outlets including Nature and Forbes."
      },
      {
        title: 'Opportunistic networking and mobile systems',
        body: 'His earlier work made foundational contributions to social-based opportunistic networking, mobility modelling, and mobile offloading. This includes influential research on social forwarding and the ThinkAir mobile cloud offloading system.'
      },
      {
        title: 'Metaverse and immersive infrastructure',
        body: 'He has advanced open AR/VR systems such as CloudAR, and has explored gesture interaction, privacy, accessibility, and large-scale spatial computing. His work on MetaHKUST connected HKUST and HKUST(GZ) in a shared immersive learning environment across campuses.'
      },
      {
        title: 'XR + AI creative practice',
        body: 'In 2025, he led and curated SURREALITY, a large-scale XR and AI art exhibition that integrated spatial computing, real-time rendering, and campus-scale storytelling, bringing metaverse ideas into a public cultural setting.'
      }
    ]
  },
  zh: {
    pageLabel: '负责人',
    documentTitle: '许彬教授 | MC2',
    profileAlt: '许彬 Pan Hui 教授',
    title: '许彬 Pan Hui 教授',
    subtitle: 'MC2 主任，港科大（广州）计算媒体与艺术学域讲席教授兼代理学域主任',
    meta: '英国皇家工程院国际院士、欧洲科学院院士、IEEE Fellow、ACM 杰出科学家。',
    homepageLabel: '个人主页',
    sections: [
      {
        title: '个人简介',
        paragraphs: [
          '许彬（Pan Hui）教授现任香港科技大学（广州）计算媒体与艺术学域讲席教授及代理学域主任、元宇宙与计算创意中心（MC2）主任；同时任香港科技大学新兴跨学科领域讲席教授、HKUST-DT 系统与媒体实验室（SyMLab）主任，以及赫尔辛基大学计算机系教授。此前，他曾担任港科大（广州）信息枢纽副院长，也是香港科技大学计算机科学与工程学系终身教职成员之一。',
          '他毕业于香港大学与剑桥大学，长期活跃于移动计算、计算机网络、增强现实、数据科学与人机交互等交叉领域。由于在移动计算和网络方向的持续贡献，他获选英国皇家工程院国际院士、欧洲科学院院士、IEEE Fellow 与 ACM 杰出科学家。截至 2026 年，他的 Google Scholar 引用已接近 40000 次。'
        ]
      },
      {
        title: '研究脉络',
        paragraphs: [
          '许教授的研究主线可以概括为“数据驱动系统设计”与“沉浸式人机数据交互”。一方面，他通过大规模实证数据理解人类行为与社会技术系统，并进一步设计面向移动设备、可穿戴设备、云边协同环境的高性能软件与网络系统；另一方面，他持续推进 AR/VR/XR、空间计算与元宇宙基础设施研究，探索数字内容、社会交互与智能系统如何更自然地进入真实空间。'
        ]
      }
    ],
    projectsTitle: '代表性方向与项目',
    projects: [
      {
        title: 'AI 赋能教育',
        body: '2024 年，他带领团队推出亚洲首批 AI 讲师项目，将 AI 生成数字人引入教学场景，系统探索沉浸式课堂中的拟真度、信任感与教学效果之间的关系，相关成果随后受到 Nature 与 Forbes 等媒体关注。'
      },
      {
        title: '机会网络与移动系统',
        body: '他在社会机会通信、移动行为建模与移动卸载等方向做出了基础性贡献，代表性工作包括 social-based forwarding 研究以及 ThinkAir 移动云卸载系统。'
      },
      {
        title: '元宇宙与沉浸式基础设施',
        body: '他推动了 CloudAR 等开放 AR/VR 系统的发展，并持续研究手势交互、隐私保护、辅助应用和大规模空间计算。MetaHKUST 项目则将港科大与港科大（广州）连接为跨校区共享的沉浸式学习环境。'
      },
      {
        title: 'XR 与 AI 创意实践',
        body: '2025 年，他主导并策划了 SURREALITY 大型 XR 与 AI 艺术展，通过空间计算、实时渲染与校园级叙事，把元宇宙相关理念带入公共文化与艺术现场。'
      }
    ]
  }
};

function useLeaderPageBodyClass(): void {
  useEffect(() => {
    document.body.classList.add('leader-page-on');
    return () => document.body.classList.remove('leader-page-on');
  }, []);
}

export default function LeaderPage({ locale = 'en' }: LeaderPageProps): JSX.Element {
  useLeaderPageBodyClass();
  const isZh = locale === 'zh-CN';
  const content = useMemo(() => (isZh ? leaderContent.zh : leaderContent.en), [isZh]);

  useEffect(() => {
    document.title = content.documentTitle;
  }, [content.documentTitle]);

  return (
    <SitePageShell className='leader-page-shell' ariaLabel={content.pageLabel} locale={locale} activeRoute='leader'>
      <div className='leader-wrap'>
        <section className='leader-card'>
          <img src='/images/Team_Profile_Pic/2.jpg?v=20260412-2' alt={content.profileAlt} />
          <div>
            <h1 className='leader-title'>{content.title}</h1>
            <p className='leader-subtitle'>{content.subtitle}</p>
            <p className='leader-meta'>{content.meta}</p>
            <a className='leader-homepage' href='https://panhui.people.ust.hk/' target='_blank' rel='noopener noreferrer'>{content.homepageLabel}</a>
          </div>
        </section>

        <section className='leader-section'>
          {content.sections.map((section) => (
            <React.Fragment key={section.title}>
              <h3>{section.title}</h3>
              {section.paragraphs.map((paragraph) => <p key={paragraph}>{paragraph}</p>)}
            </React.Fragment>
          ))}

          <h3>{content.projectsTitle}</h3>
          <ul className='leader-projects'>
            {content.projects.map((project) => (
              <li key={project.title}>
                <strong>{project.title}</strong>
                <span>{project.body}</span>
              </li>
            ))}
          </ul>
        </section>
      </div>
    </SitePageShell>
  );
}
