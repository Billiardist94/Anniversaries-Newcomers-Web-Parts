export class WeatherInfo{
    public locationRef: string;
    public capture: string;
    public temperature: number;
    public realFeel: number;
    public icon: number;
    public temperatureLo?: number;
    public temperatureHi?: number;
    public utcOffSet?: string;
}