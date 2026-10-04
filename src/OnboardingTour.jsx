import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'

const STORAGE_KEY = 'veritatem-tour-seen'

function getStepTargets() {
  return {
    dashboard: { selector: '#dashboardPage', title: 'Home / Patient Data', description: 'This is the central overview of the currently selected patient.' },
    patientCard: { selector: '.patientBadge', title: 'Patient Card', description: 'This card shows the currently selected patient and allows you to switch to another patient.' },
    sidebar: { selector: '.sidebar', title: 'Navigation / Burger Menu', description: 'The left navigation contains the main areas of the app. The V icon toggles between light and dark mode.' },
    project: { selector: '.icon:last-of-type', title: 'Project / Dataset Page', description: 'This page gives access to datasets, documents, and project information.' },
  }
}

function OnboardingTour({ enabled, onComplete, onSkip, onStepChange }) {
  const [stepIndex, setStepIndex] = useState(0)
  const [isVisible, setIsVisible] = useState(false)
  const [targetRect, setTargetRect] = useState(null)
  const [mounted, setMounted] = useState(false)
  const lastTargetRef = useRef(null)

  const steps = useMemo(
    () => [
      {
        key: 'dashboard',
        title: 'Home / Patient Data',
        description: 'The home screen shows the central overview of the currently selected patient.',
      },
      {
        key: 'patientCard',
        title: 'Patient Card in the Top Right',
        description: 'The patient card shows the currently selected patient and allows you to switch to another patient.',
      },
      {
        key: 'sidebar',
        title: 'Navigation / Burger Menu',
        description: 'The left navigation contains the main areas of the app. The V icon toggles between light and dark mode.',
      },
      {
        key: 'project',
        title: 'Project / Dataset Page',
        description: 'The final icon in the navigation leads to datasets, documents, and project information.',
      },
    ],
    [],
  )

  useEffect(() => {
    if (!enabled) {
      setIsVisible(false)
      return
    }

    setMounted(true)
    setIsVisible(true)
    setStepIndex(0)
  }, [enabled])

  useEffect(() => {
    onStepChange?.(stepIndex)
  }, [onStepChange, stepIndex])

  useLayoutEffect(() => {
    if (!isVisible) return undefined

    let frameId
    let resizeObserver

    const updateTarget = () => {
      const currentStep = steps[stepIndex]
      const target = getStepTargets()[currentStep.key]

      if (!target) {
        setTargetRect(null)
        return
      }

      const element = document.querySelector(target.selector)

      if (!element) {
        setTargetRect(null)
        return
      }

      const rect = element.getBoundingClientRect()
      const nextRect = {
        top: rect.top,
        left: rect.left,
        width: rect.width,
        height: rect.height,
      }

      setTargetRect((previous) => {
        const hasChanged = !previous || previous.top !== nextRect.top || previous.left !== nextRect.left || previous.width !== nextRect.width || previous.height !== nextRect.height
        return hasChanged ? nextRect : previous
      })
      lastTargetRef.current = element
    }

    const scheduleMeasure = () => {
      setTargetRect(null)
      frameId = window.requestAnimationFrame(() => {
        updateTarget()
      })
    }

    scheduleMeasure()

    if (typeof ResizeObserver !== 'undefined') {
      const element = document.querySelector(getStepTargets()[steps[stepIndex].key].selector)
      if (element) {
        resizeObserver = new ResizeObserver(() => {
          updateTarget()
        })
        resizeObserver.observe(element)
      }
    }

    window.addEventListener('resize', scheduleMeasure)
    window.addEventListener('scroll', scheduleMeasure, true)

    return () => {
      window.cancelAnimationFrame(frameId)
      window.removeEventListener('resize', scheduleMeasure)
      window.removeEventListener('scroll', scheduleMeasure, true)
      resizeObserver?.disconnect()
    }
  }, [isVisible, stepIndex, steps])

  const currentStep = steps[stepIndex]

  const completeTour = () => {
    localStorage.setItem(STORAGE_KEY, 'true')
    setIsVisible(false)
    onComplete?.()
  }

  const skipTour = () => {
    localStorage.setItem(STORAGE_KEY, 'true')
    setIsVisible(false)
    onSkip?.()
  }

  const nextStep = () => {
    if (stepIndex < steps.length - 1) {
      setStepIndex((previous) => previous + 1)
      return
    }

    completeTour()
  }

  const previousStep = () => {
    if (stepIndex > 0) {
      setStepIndex((previous) => previous - 1)
    }
  }

  if (!mounted || !enabled || !isVisible || !currentStep) return null

  const overlayStyle = targetRect
    ? {
        '--cutout-left': `${targetRect.left - 10}px`,
        '--cutout-top': `${targetRect.top - 10}px`,
        '--cutout-width': `${targetRect.width + 20}px`,
        '--cutout-height': `${targetRect.height + 20}px`,
        '--cutout-radius': '18px',
      }
    : {}

  return (
    <>
      {targetRect && (
        <div className="tour-overlay" style={overlayStyle} aria-hidden="true">
          <div className="tour-overlay-panel tour-overlay-top" />
          <div className="tour-overlay-panel tour-overlay-bottom" />
          <div className="tour-overlay-panel tour-overlay-left" />
          <div className="tour-overlay-panel tour-overlay-right" />
        </div>
      )}
      {targetRect && (
        <div
          className="tour-highlight"
          style={{
            top: targetRect.top - 10,
            left: targetRect.left - 10,
            width: targetRect.width + 20,
            height: targetRect.height + 20,
            borderRadius: 'var(--cutout-radius, 18px)',
          }}
        />
      )}

      <div className="tour-dialog" role="dialog" aria-live="polite">
        <div className="tour-header">
          <span className="tour-step">{stepIndex + 1} / {steps.length}</span>
          <button type="button" className="tour-close" onClick={skipTour} aria-label="Skip tour">
            ✕
          </button>
        </div>

        <h3>{currentStep.title}</h3>
        <p>{currentStep.description}</p>

        <div className="tour-actions">
          <button type="button" className="tour-secondary" onClick={previousStep} disabled={stepIndex === 0}>
            Back
          </button>
          <button type="button" className="tour-primary" onClick={nextStep}>
            {stepIndex === steps.length - 1 ? 'Finish' : 'Next'}
          </button>
        </div>
      </div>
    </>
  )
}

export { STORAGE_KEY }
export default OnboardingTour
