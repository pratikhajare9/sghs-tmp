import * as React from 'react';
import styles from './Accordion.module.scss';
import { RichText } from '@pnp/spfx-controls-react/lib/RichText';
import { ChevronDown24Regular, Add20Regular, Delete20Regular } from '@fluentui/react-icons';
import { IAccordionItem } from '../components/IAccordionProps';

interface IProps {
  items: IAccordionItem[];
  mode: string;
  isEditMode: boolean;
  onUpdate: (items: IAccordionItem[]) => void;
  headerLevel?: string; // H1, H2, H3, etc.
  fontSize?: string; // 12px, 14px, 16px, etc.
}

export const Accordion: React.FC<IProps> = ({ items, mode, isEditMode, onUpdate, headerLevel = 'h2', fontSize = '1rem' }) => {

  const toggleItem = (id: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const updated = items.map(item => {
      // In edit mode, always use single-open behavior
      if (isEditMode || mode === 'single') {
        return { ...item, isOpen: item.id === id ? !item.isOpen : false };
      }
      return item.id === id ? { ...item, isOpen: !item.isOpen } : item;
    });
    onUpdate([...updated]);
  };

  const updateTitle = (id: string, newTitle: string) => {
    const updated = items.map(item =>
      item.id === id ? { ...item, title: newTitle } : item
    );
    onUpdate([...updated]);
  };

  const updateContent = (id: string, newContent: string) => {
    const updated = items.map(item =>
      item.id === id ? { ...item, content: newContent } : item
    );
    onUpdate([...updated]);
  };

  const addItem = () => {
    const newId = `item-${Date.now()}`;
    const newItem: IAccordionItem = {
      id: newId,
      title: `Accordion Section ${items.length + 1}`,
      content: '',
      isOpen: false
    };
    const updatedItems = [...items, newItem];
    onUpdate(updatedItems);
  };

  const deleteItem = (id: string) => {
    const updated = items.filter(item => item.id !== id);
    onUpdate([...updated]);
  };

  return (
    <div className={styles.accordion}>
      {items && items.map(item => (
        <div key={item.id} className={`${styles.item} ${isEditMode && item.isOpen ? styles.itemHighlight : ''}`}>
          <button
            className={styles.header}
            onClick={(e) => toggleItem(item.id, e)}
            aria-expanded={item.isOpen}
            type="button"
          >
            <div style={{ display: 'flex', alignItems: 'center', flex: 1, gap: '8px' }}>
              {isEditMode ? (
                <input
                  type="text"
                  value={item.title}
                  onChange={(e) => updateTitle(item.id, e.target.value)}
                  onClick={(e) => e.stopPropagation()}
                  style={{
                    flex: 1,
                    padding: '4px 8px',
                    border: '1px solid #00678f',
                    borderRadius: '2px',
                    fontSize: fontSize,
                    fontFamily: 'inherit',
                    fontWeight: headerLevel === 'h1' ? 'bold' : headerLevel === 'h2' ? '600' : 'normal'
                  }}
                  placeholder="Header text"
                />
              ) : (
                React.createElement(
                  headerLevel as any,
                  { 
                    style: { 
                      margin: '0',
                      fontSize: fontSize,
                      fontWeight: headerLevel === 'h1' ? 'bold' : headerLevel === 'h2' ? '600' : 'normal'
                    }
                  },
                  item.title
                )
              )}
              {isEditMode && (
                <button
                  type="button"
                  onClick={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    deleteItem(item.id);
                  }}
                  style={{
                    background: 'none',
                    border: 'none',
                    color: 'var(--bodyText)',
                    cursor: 'pointer',
                    padding: '4px',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center'
                  }}
                  title="Delete section"
                >
                  <Delete20Regular />
                </button>
              )}
            </div>
            <ChevronDown24Regular
              className={`${styles.icon} ${item.isOpen ? styles.rotate : ''}`}
            />
          </button>

          <div className={`${styles.panel} ${item.isOpen ? styles.open : ''}`} style={{ visibility: item.isOpen ? 'visible' : 'hidden' }}>
            {isEditMode ? (
              <div className={styles.richTextContainer} style={{ position: 'relative', zIndex: 10 }}>
                <RichText
                  value={item.content}
                  isEditMode={isEditMode}
                  onChange={(value) => {
                    updateContent(item.id, value);
                    return value;
                  }}
                />
              </div>
            ) : (
              <div dangerouslySetInnerHTML={{ __html: item.content }} style={{ padding: '0px 10px' }} />
            )}
          </div>
          <div className={styles.divider} />
        </div>
      ))}
      
      {isEditMode && (
        <button
          type="button"
          onClick={addItem}
          style={{
            width: '100%',
            padding: '12px',
            marginTop: '16px',
            background: 'var(--buttonBackground)',
            color: 'var(--buttonText)',
            border: '1px solid var(--neutralLight)',
            borderRadius: '2px',
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            gap: '8px',
            fontSize: '1rem',
            fontFamily: 'inherit'
          }}
          title="Add new section"
        >
          <Add20Regular />
          Add Section
        </button>
      )}
    </div>
  );
};
