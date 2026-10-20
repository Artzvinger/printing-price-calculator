import { renderHook, act } from '@testing-library/react';
import { useLocalStorage } from './useLocalStorage';

describe('useLocalStorage', () => {
    beforeEach(() => {
        localStorage.clear();
    });

    it('возвращает начальное значение при пустом localStorage', () => {
        const { result } = renderHook(() =>
            useLocalStorage('test-key', { value: 1 })
        );

        expect(result.current[0]).toEqual({ value: 1 });
    });

    it('восстанавливает сохранённое значение из localStorage', () => {
        localStorage.setItem(
            'test-key',
            JSON.stringify({ value: 42 })
        );

        const { result } = renderHook(() =>
            useLocalStorage('test-key', { value: 1 })
        );

        expect(result.current[0]).toEqual({ value: 42 });
    });

    it('сохраняет новое значение в localStorage', () => {
        const { result } = renderHook(() =>
            useLocalStorage('test-key', 'initial')
        );

        act(() => {
            result.current[1]('updated');
        });

        expect(localStorage.getItem('test-key')).toBe(
            JSON.stringify('updated')
        );
    });

    it('падает к начальному значению при повреждённом JSON', () => {
        localStorage.setItem('test-key', '{not-valid-json');

        const { result } = renderHook(() =>
            useLocalStorage('test-key', 'fallback')
        );

        expect(result.current[0]).toBe('fallback');
    });
});
