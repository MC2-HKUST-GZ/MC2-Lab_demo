import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import FloatingContactMenu from '../FloatingContactMenu';
import { useSectionScrollSpy } from '../../hooks/useSectionScrollSpy';
import type { AlumniGroup, MemberProfile, MemberRole, PeoplePageContent } from '../../types/legacyPages';

interface PeopleContentProps {
  content: PeoplePageContent;
}

const ROLE_SORT_ORDER: Record<string, number> = {
  leader: 0,
  chair: 0,
  postdoc: 1,
  admin: 2,
  phd: 3,
  mphil: 4,
  ra: 5,
  other: 6
};

function roleRank(member: MemberProfile): number {
  return ROLE_SORT_ORDER[member.role] ?? 99;
}

function cohortYearValue(member: MemberProfile): number {
  return Number.isFinite(member.cohortYear) ? Number(member.cohortYear) : 9999;
}

function formatFrontPosition(member: MemberProfile): string {
  if (member.cohortYear && member.cohortTerm) return `${member.positionTitle} | ${member.cohortTerm} ${member.cohortYear}`;
  return member.positionTitle;
}

function sortMembers(members: MemberProfile[]): MemberProfile[] {
  return [...members].sort((a, b) => {
    const roleA = roleRank(a);
    const roleB = roleRank(b);
    if (roleA !== roleB) return roleA - roleB;

    const yearA = cohortYearValue(a);
    const yearB = cohortYearValue(b);
    if (yearA !== yearB) return yearA - yearB;

    return a.name.localeCompare(b.name, 'en', { sensitivity: 'base' });
  });
}

function matchesMember(member: MemberProfile, query: string, role: MemberRole): boolean {
  const rolePass = role === 'all' || member.role === role;
  if (!rolePass) return false;
  if (!query) return true;

  const haystack = [
    member.name,
    member.positionTitle,
    member.researchDirection,
    member.role,
    member.cohortTerm,
    member.cohortYear
  ].join(' ').toLowerCase();

  return haystack.includes(query);
}

function MemberCard({ fallbackResearch, member }: { fallbackResearch: string; member: MemberProfile }): JSX.Element {
  const [isVisible, setIsVisible] = useState(false);
  const [photo, setPhoto] = useState(member.photo);
  const photoClassName = member.key === 'pan-hui' ? 'card__img card__img--pan-hui' : 'card__img';
  const research = member.researchDirection || fallbackResearch;

  return (
    <li className='member-card-item'>
      <article className='card' data-visible={isVisible ? 'true' : 'false'}>
        <div className='card__front flow-content'>
          <img
            className={`${photoClassName} mx-auto`}
            src={photo}
            alt={`${member.name} profile photo`}
            onError={() => setPhoto('/images/Team_Profile_Pic/profile_icon.jpg')}
          />
          <div className='flow-content' data-spacing='sm'>
            <p className='card__name'>{member.name}</p>
            <p className='card__position'>{formatFrontPosition(member)}</p>
          </div>
        </div>

        <div className='card__back flow-content'>
          <p className='card__name'>{member.name}</p>
          <q>{research}</q>
          <ul role='list' className='card__social flex-group'>
            {member.website ? (
              <li>
                <a href={member.website} target='_blank' rel='noopener noreferrer'>
                  <i className='fa fa-id-card fa-xl' style={{ paddingTop: '1em' }}></i>
                </a>
              </li>
            ) : null}
          </ul>
        </div>

        <footer className='card__footer'>
          <button
            data-card-controller
            className='card__toggle'
            aria-label={`Toggle card details for ${member.name}`}
            onClick={() => setIsVisible((visible) => !visible)}
          >
            <i className='fa-solid fa-plus card__toggle-icon'></i>
          </button>
        </footer>
      </article>
    </li>
  );
}

function MemberExplorer({ content }: PeopleContentProps): JSX.Element {
  const [query, setQuery] = useState('');
  const [activeRole, setActiveRole] = useState<MemberRole>('all');

  const visibleMembers = useMemo(() => {
    const normalizedQuery = query.trim().toLowerCase();
    return sortMembers(content.members).filter((member) => matchesMember(member, normalizedQuery, activeRole));
  }, [activeRole, content.members, query]);

  return (
    <>
      <div className='member-explorer' aria-label={content.searchAria}>
        <label className='member-search-shell' htmlFor='memberSearchInput'>
          <i className='fa-solid fa-magnifying-glass member-search-icon' aria-hidden='true'></i>
          <input
            id='memberSearchInput'
            className='member-search-input'
            type='search'
            placeholder={content.searchPlaceholder}
            value={query}
            onChange={(event) => setQuery(event.target.value)}
          />
        </label>
        <div className='member-role-chips' id='memberRoleChips' role='group' aria-label={content.roleFilterAria}>
          {content.roleChips.map((chip) => (
            <button
              type='button'
              className={activeRole === chip.role ? 'member-chip is-active' : 'member-chip'}
              data-role-filter={chip.role}
              key={chip.role}
              onClick={() => setActiveRole(chip.role)}
            >
              {chip.label}
            </button>
          ))}
        </div>
      </div>

      <section id='Current'>
        <ul role='list' className='grid' id='membersGrid'>
          {visibleMembers.length ? visibleMembers.map((member) => (
            <MemberCard fallbackResearch={content.fallbackResearch} key={member.key} member={member} />
          )) : <li className='member-empty'>{content.emptyText}</li>}
        </ul>
      </section>
    </>
  );
}

function AlumniSections({ groups }: { groups: AlumniGroup[] }): JSX.Element {
  return (
    <section id='Graduated'>
      <h1 className='section-title' style={{ fontSize: '3em', paddingTop: '1em' }}>Graduated / Alumni</h1>
      {groups.map((group) => (
        <React.Fragment key={group.heading}>
          <h3 style={{ color: '#fff', paddingBottom: '.5em' }}>{group.heading}</h3>
          <ul>
            {group.entries.map((entry, index) => (
              <li key={`${group.heading}-${entry.name}-${index}`}>
                <span className='bold-name'>{entry.name}</span> {entry.detail}
              </li>
            ))}
          </ul>
        </React.Fragment>
      ))}
    </section>
  );
}

export default function PeopleContent({ content }: PeopleContentProps): JSX.Element {
  const pageRef = useRef<HTMLDivElement>(null);
  const lastScrollTopRef = useRef(0);
  const programmaticScrollTargetRef = useRef<number | null>(null);
  const programmaticScrollDeadlineRef = useRef(0);
  const scrollFrameRef = useRef<number | null>(null);
  const [activeSection, setActiveSection] = useState('Current');
  const [isSubnavHidden, setIsSubnavHidden] = useState(false);
  const sectionIds = useMemo(() => ['Current', 'Graduated'], []);

  useSectionScrollSpy(pageRef, sectionIds, setActiveSection, 190);

  useEffect(() => {
    const scrollContainer = pageRef.current;
    if (!scrollContainer) return undefined;

    const updateSubnavState = () => {
      const nextScrollTop = scrollContainer.scrollTop;
      const delta = nextScrollTop - lastScrollTopRef.current;
      const programmaticTarget = programmaticScrollTargetRef.current;

      if (programmaticTarget !== null) {
        setIsSubnavHidden(false);
        if (Math.abs(nextScrollTop - programmaticTarget) < 4 || Date.now() > programmaticScrollDeadlineRef.current) {
          programmaticScrollTargetRef.current = null;
        }
        lastScrollTopRef.current = nextScrollTop;
        scrollFrameRef.current = null;
        return;
      }

      if (nextScrollTop < 120) {
        setIsSubnavHidden(false);
      } else if (delta > 12) {
        setIsSubnavHidden(true);
      } else if (delta < -10) {
        setIsSubnavHidden(false);
      }

      lastScrollTopRef.current = nextScrollTop;
      scrollFrameRef.current = null;
    };

    const handleScroll = () => {
      if (scrollFrameRef.current !== null) return;
      scrollFrameRef.current = window.requestAnimationFrame(updateSubnavState);
    };

    scrollContainer.addEventListener('scroll', handleScroll, { passive: true });

    return () => {
      scrollContainer.removeEventListener('scroll', handleScroll);
      if (scrollFrameRef.current !== null) window.cancelAnimationFrame(scrollFrameRef.current);
    };
  }, []);

  const selectSection = useCallback((sectionId: string) => {
    setActiveSection(sectionId);
    setIsSubnavHidden(false);
    const scrollContainer = pageRef.current;
    const section = scrollContainer?.querySelector<HTMLElement>(`#${sectionId}`);
    if (!scrollContainer || !section) return;
    const containerRect = scrollContainer.getBoundingClientRect();
    const sectionTop = section.getBoundingClientRect().top - containerRect.top + scrollContainer.scrollTop;
    const stickyOffset = window.matchMedia('(max-width: 980px)').matches ? 172 : 90;
    const targetTop = Math.max(0, sectionTop - stickyOffset);
    programmaticScrollTargetRef.current = targetTop;
    programmaticScrollDeadlineRef.current = Date.now() + 4000;
    scrollContainer.scrollTo({ top: targetTop, behavior: 'smooth' });
  }, []);

  return (
    <>
      <div className='people-content' data-subnav-state={isSubnavHidden ? 'hidden' : 'visible'} ref={pageRef}>
        <section>
          <div className='publication-nav' aria-label={content.pageHeading}>
            <div className='container'>
              <ul className='yr__navs'>
                <li className={activeSection === 'Current' ? 'yr__nav yr2024 active' : 'yr__nav yr2024'}>
                  <a href='#Current' onClick={(event) => { event.preventDefault(); selectSection('Current'); }}>{content.currentLabel}</a>
                </li>
                <li className={activeSection === 'Graduated' ? 'yr__nav yr2023 active' : 'yr__nav yr2023'}>
                  <a href='#Graduated' onClick={(event) => { event.preventDefault(); selectSection('Graduated'); }}>{content.alumniLabel}</a>
                </li>
              </ul>
            </div>
          </div>

          <div className='container flow-content'>
            <h1 className='section-title' style={{ fontSize: '6em' }}>{content.pageHeading}</h1>
            <MemberExplorer content={content} />
            <AlumniSections groups={content.alumniGroups} />
          </div>
        </section>
      </div>
      <FloatingContactMenu />
    </>
  );
}
