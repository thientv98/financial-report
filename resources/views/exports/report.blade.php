<table>
    <thead>
        <tr>
            <th>Code: {{$code}}</th>
        </tr>
        <tr>
            @foreach ($head as $item)
                <th>{{ $item }}</th>
            @endforeach
        </tr>
    </thead>
    <tbody>
        @foreach ($body as $items)
            <tr>
                @foreach ($items as $item)
                    <td>{{ $item }}</td>
                @endforeach
            </tr>
        @endforeach
    </tbody>
</table>