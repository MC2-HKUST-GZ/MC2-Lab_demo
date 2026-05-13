import { useEffect } from 'react';

export function useSectionScrollSpy(
  scrollContainerRef: React.RefObject<HTMLElement>,
  sectionIds: string[],
  onActiveSectionChange: (sectionId: string) => void,
  offset = 140
): void {
  useEffect(() => {
    const scrollContainer = scrollContainerRef.current;
    if (!scrollContainer || sectionIds.length === 0) return undefined;

    const syncActiveSection = () => {
      const sections = sectionIds
        .map((sectionId) => scrollContainer.querySelector<HTMLElement>(`#${sectionId}`))
        .filter((section): section is HTMLElement => Boolean(section));
      if (!sections.length) return;

      const containerRect = scrollContainer.getBoundingClientRect();
      const triggerLine = scrollContainer.scrollTop + offset;
      const isAtBottom = scrollContainer.scrollTop + scrollContainer.clientHeight >= scrollContainer.scrollHeight - 2;
      const visibleBottomLine = scrollContainer.scrollTop + scrollContainer.clientHeight - offset;
      let activeSection = sections[0];

      sections.forEach((section) => {
        const sectionTop = section.getBoundingClientRect().top - containerRect.top + scrollContainer.scrollTop;
        if (sectionTop <= (isAtBottom ? visibleBottomLine : triggerLine)) activeSection = section;
      });

      onActiveSectionChange(activeSection.id);
    };

    scrollContainer.addEventListener('scroll', syncActiveSection, { passive: true });
    window.addEventListener('resize', syncActiveSection);
    syncActiveSection();

    return () => {
      scrollContainer.removeEventListener('scroll', syncActiveSection);
      window.removeEventListener('resize', syncActiveSection);
    };
  }, [offset, onActiveSectionChange, scrollContainerRef, sectionIds]);
}
