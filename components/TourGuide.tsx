import React, { useEffect, useState, useRef } from 'react';
import { createPortal } from 'react-dom';
import { X, ChevronRight, ChevronLeft } from 'lucide-react';

export interface TourStep {
  target: string; // CSS Selector (#id or .class)
  title: string;
  content: string;
  position?: 'top' | 'bottom' | 'left' | 'right' | 'center';
  additionalTargets?: string[]; // New: support multiple highlights
  onEnter?: () => void; // Trigger when step starts
  onLeave?: () => void; // Trigger when step ends
}

interface TourGuideProps {
  steps: TourStep[];
  isOpen: boolean;
  onClose: () => void;
}

interface Rect {
  top: number;
  left: number;
  width: number;
  height: number;
}

const TourGuide: React.FC<TourGuideProps> = ({ steps, isOpen, onClose }) => {
  const [currentStep, setCurrentStep] = useState(0);
  const [targetRects, setTargetRects] = useState<Rect[]>([]);
  const [popoverStyle, setPopoverStyle] = useState<React.CSSProperties>({});
  const popoverRef = useRef<HTMLDivElement>(null);

  // Reset step when opening
  useEffect(() => {
    if (isOpen) {
      setCurrentStep(0);
    } else {
      setTargetRects([]);
    }
  }, [isOpen]);

  // Handle lifecycle and positioning
  useEffect(() => {
    if (!isOpen) return;

    const step = steps[currentStep];
    if (!step) return;

    // Trigger onEnter
    if (step.onEnter) {
        step.onEnter();
    }

    const updatePosition = () => {
      if (step.position === 'center') {
        setTargetRects([]);
        setPopoverStyle({});
        return;
      }

      const targets = [step.target, ...(step.additionalTargets || [])];
      const newRects: Rect[] = [];
      let mainRect: DOMRect | null = null;

      targets.forEach((selector, index) => {
          const element = document.querySelector(selector);
          if (element) {
             const rect = element.getBoundingClientRect();
             const padding = 6;
             newRects.push({
                 top: rect.top - padding,
                 left: rect.left - padding,
                 width: rect.width + (padding * 2),
                 height: rect.height + (padding * 2)
             });
             
             // First target is the main one for popover positioning
             if (index === 0) {
                 mainRect = rect;
                 element.scrollIntoView({ behavior: 'smooth', block: 'center' });
             }
          }
      });

      setTargetRects(newRects);

      // Popover Positioning Logic based on Main Target
      if (mainRect && popoverRef.current) {
            const popRect = popoverRef.current.getBoundingClientRect();
            const padding = 6;
            const margin = 15;
            let popTop = 0;
            let popLeft = 0;

            const rect = mainRect; // Use the raw DOMRect for calculation logic

            switch (step.position) {
                case 'right':
                    popLeft = rect.right + margin + padding;
                    popTop = rect.top + (rect.height / 2) - (popRect.height / 2);
                    break;
                case 'left':
                    popLeft = rect.left - popRect.width - margin - padding;
                    popTop = rect.top + (rect.height / 2) - (popRect.height / 2);
                    break;
                case 'top':
                    popLeft = rect.left + (rect.width / 2) - (popRect.width / 2);
                    popTop = rect.top - popRect.height - margin - padding;
                    break;
                case 'bottom':
                default:
                    popLeft = rect.left + (rect.width / 2) - (popRect.width / 2);
                    popTop = rect.bottom + margin + padding;
                    break;
            }

            // Clamp to viewport
            // Horizontal
            if (popLeft < 10) popLeft = 10;
            if (popLeft + popRect.width > window.innerWidth - 10) popLeft = window.innerWidth - popRect.width - 10;
            
            // Vertical
            if (popTop < 10) popTop = 10;
            if (popTop + popRect.height > window.innerHeight - 10) popTop = window.innerHeight - popRect.height - 10;

            setPopoverStyle({
                position: 'absolute',
                top: popTop,
                left: popLeft,
            });
        }
    };

    // Initial updates with delays to account for rendering/animations
    updatePosition();
    const t1 = setTimeout(updatePosition, 100);
    const t2 = setTimeout(updatePosition, 300); // Retry for modals
    const t3 = setTimeout(updatePosition, 600); 

    window.addEventListener('resize', updatePosition);

    return () => {
      window.removeEventListener('resize', updatePosition);
      clearTimeout(t1);
      clearTimeout(t2);
      clearTimeout(t3);
      if (step.onLeave) {
          step.onLeave();
      }
    };
  }, [currentStep, isOpen, steps]);

  if (!isOpen) return null;

  const step = steps[currentStep];
  const isLastStep = currentStep === steps.length - 1;

  const handleNext = () => {
    if (isLastStep) {
      onClose();
    } else {
      setCurrentStep(prev => prev + 1);
    }
  };

  const handlePrev = () => {
    if (currentStep > 0) {
      setCurrentStep(prev => prev - 1);
    }
  };

  return createPortal(
    <div className="fixed inset-0 z-[9999] overflow-hidden">
      {/* SVG Backdrop with Mask for Holes */}
      {targetRects.length > 0 && (
          <svg className="absolute inset-0 w-full h-full pointer-events-none transition-opacity duration-500">
            <defs>
                <mask id="tour-mask">
                    <rect x="0" y="0" width="100%" height="100%" fill="white" />
                    {targetRects.map((rect, i) => (
                        <rect 
                            key={i}
                            x={rect.left} 
                            y={rect.top} 
                            width={rect.width} 
                            height={rect.height} 
                            fill="black" 
                            rx="8" // Rounded corners for mask
                        />
                    ))}
                </mask>
            </defs>
            <rect 
                x="0" 
                y="0" 
                width="100%" 
                height="100%" 
                fill="rgba(0,0,0,0.7)" 
                mask="url(#tour-mask)" 
            />
          </svg>
      )}

      {/* Fallback full backdrop if no targets (e.g. center step) */}
      {targetRects.length === 0 && (
          <div className="absolute inset-0 bg-black/70 transition-opacity duration-500" />
      )}

      {/* Highlight Borders */}
      {targetRects.map((rect, i) => (
        <div 
          key={i}
          className="absolute rounded-lg transition-all duration-300 ease-in-out pointer-events-none border-4 border-yellow-400 animate-pulse shadow-[0_0_20px_rgba(250,204,21,0.5)]"
          style={{
            top: rect.top,
            left: rect.left,
            width: rect.width,
            height: rect.height,
          }}
        />
      ))}

      {/* Popover Card */}
      <div className="absolute w-full h-full pointer-events-none top-0 left-0">
          <div 
            ref={popoverRef}
            className={`
                pointer-events-auto bg-white dark:bg-slate-800 rounded-xl shadow-2xl p-6 max-w-sm w-full border border-gray-100 dark:border-slate-600 animate-in fade-in zoom-in duration-300
                ${targetRects.length === 0 ? 'absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2' : ''}
            `}
            style={targetRects.length > 0 ? popoverStyle : {}}
          >
            <div className="flex justify-between items-start mb-4">
                <div className="flex items-center gap-3">
                    <span className="bg-blue-600 text-white text-xs font-bold px-3 py-1 rounded-full whitespace-nowrap flex items-center justify-center shadow-sm">
                        {currentStep + 1} / {steps.length}
                    </span>
                    <h3 className="text-lg font-bold text-gray-800 dark:text-white leading-tight">{step.title}</h3>
                </div>
                <button onClick={onClose} className="text-gray-400 hover:text-gray-600 dark:text-gray-500 dark:hover:text-gray-300 shrink-0 ml-2">
                    <X size={20} />
                </button>
            </div>
            
            <p className="text-gray-600 dark:text-gray-300 text-sm mb-6 leading-relaxed">
                {step.content}
            </p>

            <div className="flex justify-between items-center pt-2 border-t border-gray-100 dark:border-slate-700">
                <button 
                    onClick={onClose}
                    className="text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200 text-sm font-medium"
                >
                    Pular
                </button>

                <div className="flex gap-2">
                    {currentStep > 0 && (
                        <button 
                            onClick={handlePrev}
                            className="p-2 rounded-lg hover:bg-gray-100 dark:hover:bg-slate-700 text-gray-600 dark:text-gray-300 transition-colors"
                        >
                            <ChevronLeft size={20} />
                        </button>
                    )}
                    <button 
                        onClick={handleNext}
                        className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-bold text-sm transition-all shadow-md hover:shadow-lg transform hover:-translate-y-0.5"
                    >
                        {isLastStep ? 'Concluir' : 'Próximo'}
                        {!isLastStep && <ChevronRight size={16} />}
                    </button>
                </div>
            </div>
          </div>
      </div>
    </div>,
    document.body
  );
};

export default TourGuide;